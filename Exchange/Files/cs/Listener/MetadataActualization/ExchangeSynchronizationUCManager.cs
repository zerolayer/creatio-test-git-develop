namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Sync;

	#region Class: ExchangeSynchronizationUCManager

	/// <summary>
	/// Class provides users that has enabled exchange calendars synchronization list.
	/// </summary>
	public class ExchangeSynchronizationUCManager : ISynchronizationUCManager
	{

		#region Properties: Public

		/// <summary>
		/// <see cref="ISynchronizationUCManager.SynchronizationControllerName"/>
		/// </summary>
		public virtual string SynchronizationControllerName { get; } = "SysSyncMetaDataActualizer";

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns related to <paramref name="entity"/> parent record unique identifier.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <returns>Related to <paramref name="entity"/> activity unique identifier.</returns>
		protected virtual Guid GetRelatedEntityId(Entity entity) {
			var columnName = GetRelatedEntityColumnName(entity);
			return columnName.IsNotNullOrEmpty() && entity.IsColumnValueLoaded(columnName)
					? entity.GetTypedColumnValue<Guid>(columnName)
					: Guid.Empty;
		}

		/// <summary>
		/// Returns related to <paramref name="entity"/> parent column name.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <returns>Related to <paramref name="entity"/> activity unique identifier.</returns>
		protected virtual string GetRelatedEntityColumnName(Entity entity) {
			return string.Empty;
		}

		/// <summary>
		/// Searches users that can synchronize <paramref name="entityId"/> to exchange.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="entityId"><see cref="Entity"/> instance unique identifier.</param>
		/// <returns>Users that can synchronize <paramref name="entityId"/> to exchange.</returns>
		protected virtual Select GetUsersSelect(UserConnection userConnection, Guid entityId) {
			return null;
		}

		/// <summary>
		/// Returns list of users, that can synchronize delete action for <paramref name="entityId"/> instance.
		/// </summary>
		/// <param name="entityId"><see cref="Entity"/> instance id.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Users that can synchronize delete action of <paramref name="entityId"/>.</returns>
		protected virtual List<Guid> GetUsersForDeleteAction(Guid entityId, UserConnection userConnection) {
			var metadataSelect = new Select(userConnection).Top(1)
					.Column("Id")
				.From("SysSyncMetaData")
				.Where("LocalId").IsEqual(Column.Parameter(entityId)) as Select;
			if (metadataSelect.ExecuteScalar<Guid>().IsEmpty()) {
				return new List<Guid>();
			}
			return new List<Guid> { userConnection.CurrentUser.ContactId };
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="ISynchronizationUCManager.GetSynchronizationUsers(Entity, SyncAction, UserConnection)"/>
		/// </summary>
		public IEnumerable<Guid> GetSynchronizationUsers(Entity entity, SyncAction action, UserConnection userConnection) {
			var result = new List<Guid>();
			if (!userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata")) {
				return result;
			}
			var entityId = GetRelatedEntityId(entity);
			if (entityId.IsEmpty()) {
				return result;
			}
			if (action == SyncAction.Delete) {
				return GetUsersForDeleteAction(entityId, userConnection);
			}
			var select = GetUsersSelect(userConnection, entityId);
			if (select != null) {
				using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
					using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
						while (reader.Read()) {
							result.AddIfNotExists(reader.GetColumnValue<Guid>("CreatedById"));
						}
					}
				}
			}
			return result;
		}

		/// <summary>
		/// <see cref="ISynchronizationUCManager.GetParentRelationColumnValue(Entity, UserConnection)"/>
		/// </summary>
		public EntityColumnValue GetParentRelationColumnValue(Entity entity, UserConnection userConnection) { 
			if (!userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata")) {
				return null;
			}
			return entity.FindEntityColumnValue(GetRelatedEntityColumnName(entity));
		}

		#endregion

	}

	#endregion

}
