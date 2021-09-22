namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities.AsyncOperations;
	using Terrasoft.Core.Entities.AsyncOperations.Interfaces;
	using Terrasoft.Core.Factories;
	
	#region Class: SynchronizationControllerManager

	/// <summary>
	/// Class starts synchronization using <see cref="ISynchronizationController"/> implementations.
	/// </summary>
	public class SynchronizationControllerManager : IEntityEventAsyncOperation
	{

		#region Methods: Protected

		/// <summary>
		/// Returns related entity ModifiedOn column value. Returns <c>DateTime.UtcNow</c> if ModifiedOn column value not avaliable.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="arguments">Changed entity synchronization events.</param>
		/// <returns>ModifiedOn column value.</returns>
		protected DateTime GetModifiedOnColumnValue(UserConnection userConnection, SyncEntityEventAsyncOperationArgs arguments) {
			if (!arguments.EntityColumnValues.ContainsKey("ModifiedOn")) {
				return DateTime.UtcNow;
			}
			var modifiedOn = (DateTime)arguments.EntityColumnValues["ModifiedOn"];
			return modifiedOn.Kind == DateTimeKind.Utc ? modifiedOn : TimeZoneInfo.ConvertTimeToUtc(modifiedOn, userConnection.CurrentUser.TimeZone);
		}

		/// <summary>
		/// Creates parameters collection for synchronization controllers execution.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="arguments">Changed entity synchronization events.</param>
		/// <returns>Parameters collection for synchronization controllers execution.</returns>
		protected virtual Dictionary<string, object> GetSynchronizationParams(UserConnection userConnection, SyncEntityEventAsyncOperationArgs arguments) {
			return new Dictionary<string, object> {
				{ "EntityId", arguments.EntityId },
				{ "EntitySchemaName", arguments.EntitySchemaName },
				{ "ModifiedOn", GetModifiedOnColumnValue(userConnection, arguments) },
				{ "SyncAction", (int)arguments.Action },
				{ "UserContactId", arguments.UserContactId },
				{ "ColumnValues", arguments.EntityColumnValues }
			};
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="IEntityEventAsyncOperation.Execute"/>
		/// </summary>
		public void Execute(UserConnection userConnection, EntityEventAsyncOperationArgs arguments) {
			if (!userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata")) {
				return;
			}
			var synchronizationArgs = (SyncEntityEventAsyncOperationArgs)arguments;
			var parameters = GetSynchronizationParams(userConnection, synchronizationArgs);
			foreach (var controllerName in synchronizationArgs.Controllers) {
				var controller = ClassFactory.Get<ISynchronizationController>(controllerName);
				controller.Execute(userConnection, parameters);
			}
		}

		#endregion

	}

	#endregion

}
