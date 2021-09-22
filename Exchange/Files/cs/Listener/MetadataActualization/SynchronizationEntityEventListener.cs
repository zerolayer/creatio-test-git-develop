namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.AsyncOperations.Interfaces;
	using Terrasoft.Core.Entities.Events;
	using Terrasoft.Core.Factories;
	using Terrasoft.Sync;
	#region Class: SynchronizationEntityEventListener

	/// <summary>
	/// Class provides synchronization entities events handling.
	/// </summary>
	[EntityEventListener(SchemaName = "Activity")]
	[EntityEventListener(SchemaName = "ActivityParticipant")]
	[EntityEventListener(SchemaName = "Contact")]
	[EntityEventListener(SchemaName = "ContactCommunication")]
	[EntityEventListener(SchemaName = "ContactAddress")]
	public class SynchronizationEntityEventListener : BaseEntityEventListener
	{

		#region Methods: Private

		/// <summary>
		/// Collects users with synchronizations list, using <see cref="ISynchronizationUCManager"/> implementations.
		/// </summary>
		/// <param name="ucManagers">ISynchronizationUCManager implementation collection.</param>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <param name="action">Synchronization action for <paramref name="entity"/>.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Users with synchronizations list.</returns>
		private Dictionary<Guid, List<string>> GetUsers(IEnumerable<ISynchronizationUCManager> ucManagers, Entity entity,
				SyncAction action, UserConnection userConnection) {
			var users = new Dictionary<Guid, List<string>>();
			foreach (var ucManager in ucManagers) {
				var synchronizationUsers = ucManager.GetSynchronizationUsers(entity, action,  userConnection);
				foreach (var userName in synchronizationUsers) {
					if (!users.ContainsKey(userName)) {
						var syncControllers = new List<string> { ucManager.SynchronizationControllerName };
						users.Add(userName, syncControllers);
					} else {
						users[userName].Add(ucManager.SynchronizationControllerName);
					}
				}
			}
			return users;
		}

		/// <summary>
		/// Collects users with synchronizations list, using <see cref="ISynchronizationUCManager"/> implementations.
		/// </summary>
		/// <param name="ucManagers">ISynchronizationUCManager implementation collection.</param>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <param name="modifiedColumnValues">Synchronization action for <paramref name="entity"/>.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		private void AddRelationColums(IEnumerable<ISynchronizationUCManager> ucManagers, Entity entity,
				EntityColumnValueCollection modifiedColumnValues) {
			var users = new Dictionary<Guid, List<string>>();
			foreach (var ucManager in ucManagers) {
				var columnValue = ucManager.GetParentRelationColumnValue(entity, entity.UserConnection);
				if (columnValue != null && columnValue.IsLoaded && !modifiedColumnValues.Any(cv => cv.Name == columnValue.Name)) {
					modifiedColumnValues.Add(columnValue);
				}
			}
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Starts <paramref name="entity"/> synchronization action for users that has synchronization.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <param name="e"><paramref name="entity"/> event arguments instance.</param>
		/// <param name="action">Synchronization action.</param>
		protected void SyncEntity(Entity entity, EntityAfterEventArgs e, SyncAction action) {
			var userConnection = entity.UserConnection;
			if (!userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata") ||
					!ClassFactory.HasBinding(typeof(ISynchronizationUCManager))) {
				return;
			}
			if (e.ModifiedColumnValues == null) {
				e.ModifiedColumnValues = new EntityColumnValueCollection(userConnection);
			}
			var ucManagers = ClassFactory.GetAll<ISynchronizationUCManager>();
			AddRelationColums(ucManagers, entity, e.ModifiedColumnValues);
			var asyncExecutor = ClassFactory.Get<IEntityEventAsyncExecutor>(new ConstructorArgument("userConnection", userConnection));
			foreach (var user in GetUsers(ucManagers, entity, action, entity.UserConnection)) {
				var controllerNames = user.Value;
				var operationArgs = new SyncEntityEventAsyncOperationArgs(entity, e);
				operationArgs.Controllers = controllerNames;
				operationArgs.UserContactId = user.Key;
				operationArgs.Action = action;
				asyncExecutor.ExecuteAsync<SynchronizationControllerManager>(operationArgs);
			}
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnInserted"/>
		/// </summary>
		public override void OnInserted(object sender, EntityAfterEventArgs e) {
			base.OnInserted(sender, e);
			SyncEntity((Entity)sender, e, SyncAction.Create);
		}

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnUpdated"/>
		/// </summary>
		public override void OnUpdated(object sender, EntityAfterEventArgs e) {
			base.OnUpdated(sender, e);
			SyncEntity((Entity)sender, e, SyncAction.Update);
		}

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnDeleted"/>
		/// </summary>
		public override void OnDeleted(object sender, EntityAfterEventArgs e) {
			base.OnDeleted(sender, e);
			SyncEntity((Entity)sender, e, SyncAction.Delete);
		}

		#endregion

	}

	#endregion

	
}
