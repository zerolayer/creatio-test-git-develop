namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.Collections.Immutable;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Sync;

	#region Class: SysSyncMetaDataActualizer

	/// <summary>
	/// Class implementation of <see cref="ISynchronizationController"/>.
	/// </summary>
	public class SysSyncMetaDataActualizer : ISynchronizationController
	{
		#region Field: Protected

		/// <summary>
		/// Schema order.
		/// </summary>
		protected int SchemaOrder;

		/// <summary>
		/// Remote store identifier.
		/// </summary>
		protected Guid StoreId;

		/// <summary>
		/// Remote item name.
		/// </summary>
		protected string RemoteItemName;

		#endregion

		#region Methods: Private

		private void ActualizeDetailMetaData(UserConnection userConnection, MetaDataInfo metaDataInfo) {
			metaDataInfo.SetParentProperties(userConnection);
			if (metaDataInfo.ParentRemoteId.IsNullOrEmpty()) {
				return;
			}
			if (metaDataInfo.SyncAction == SyncAction.Create) {
				InsertMetaData(userConnection, metaDataInfo);
			} else {
				ActualizeMetaData(userConnection, metaDataInfo);
			}
			ActualizeMetaData(userConnection, metaDataInfo, true);
		}

		private void InsertMetaData(UserConnection userConnection, MetaDataInfo metaDataInfo) {
			var insert = new Insert(userConnection).Into("SysSyncMetaData")
				.Set("LocalId", Column.Parameter(metaDataInfo.EntityId))
				.Set("Version", Column.Parameter(metaDataInfo.ModifiedOn))
				.Set("RemoteId", Column.Parameter(metaDataInfo.ParentRemoteId))
				.Set("LocalState", Column.Parameter(SyncAction.Create))
				.Set("RemoteStoreId", Column.Parameter(StoreId))
				.Set("CreatedOn", Column.Parameter(metaDataInfo.ModifiedOn))
				.Set("ModifiedOn", Column.Parameter(metaDataInfo.ModifiedOn))
				.Set("CreatedById", Column.Parameter(metaDataInfo.UserContactId))
				.Set("ModifiedById", Column.Parameter(metaDataInfo.UserContactId))
				.Set("SyncSchemaName", Column.Parameter(metaDataInfo.EntitySchemaName))
				.Set("ModifiedInStoreId", Column.Parameter(ExchangeConsts.LocalStoreId))
				.Set("CreatedInStoreId", Column.Parameter(ExchangeConsts.LocalStoreId))
				.Set("SchemaOrder", Column.Parameter(SchemaOrder))
				.Set("RemoteItemName", Column.Parameter(RemoteItemName));
			insert.Execute();
		}

		private void ActualizeMetaData(UserConnection userConnection, MetaDataInfo metaDataInfo, bool isParentActualize = false) {
			var entityId = isParentActualize ? metaDataInfo.ParentId : metaDataInfo.EntityId;
			var entitySchemaName = isParentActualize ? metaDataInfo.ParentSchemaName : metaDataInfo.EntitySchemaName;
			var syncAction = isParentActualize ? SyncAction.Update : metaDataInfo.SyncAction;
			if (syncAction == SyncAction.Create && !IsDetailSchemaName(entitySchemaName)) {
				return;
			}
			var actualizeMetaData = new Update(userConnection, "SysSyncMetaData")
				.Set("Version", Column.Parameter(metaDataInfo.ModifiedOn))
				.Set("ModifiedOn", Column.Parameter(metaDataInfo.ModifiedOn))
				.Set("ModifiedInStoreId", Column.Parameter(ExchangeConsts.LocalStoreId))
				.Set("LocalState", Column.Parameter(syncAction))
			.Where("LocalId").IsEqual(Column.Parameter(entityId))
				.And("RemoteStoreId").IsEqual(Column.Parameter(StoreId))
				.And("SyncSchemaName").IsEqual(Column.Parameter(entitySchemaName));
			if (syncAction != SyncAction.Delete) {
				actualizeMetaData = actualizeMetaData
					.And("CreatedById").IsEqual(Column.Parameter(metaDataInfo.UserContactId)) as Update;
			}
			actualizeMetaData.Execute();
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Set additional parametrs for actualize metadatas.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected virtual void SetSyncParametrs(string entitySchemaName) {
			throw new NotImplementedException();
		}

		/// <summary>
		/// Get instance of <see cref="MetaDataInfo"/>.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected virtual MetaDataInfo GetMetaDataInfo(IDictionary<string, object> parameters) {
			throw new NotImplementedException();
		}

		/// <summary>
		/// Indicates detals schema.
		/// </summary>
		/// <param name="syncSchemaName">Synchronization schema name.</param>
		/// <returns>True if schema is detail, otherwise false.</returns>
		protected virtual bool IsDetailSchemaName(string syncSchemaName) {
			throw new NotImplementedException();
		}
		#endregion

		#region Methods: Public

		/// <summary>
		/// Runs <see cref="ISynchronizationController"/> implementations.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Entity synchronization parameters.</param>
		public void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			if (!userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata")) {
				return;
			}
			var metaDataInfo = GetMetaDataInfo(parameters);
			SetSyncParametrs(metaDataInfo.EntitySchemaName);
			if (IsDetailSchemaName(metaDataInfo.EntitySchemaName)) {
				ActualizeDetailMetaData(userConnection, metaDataInfo);
			} else {
				ActualizeMetaData(userConnection, metaDataInfo);
			}
		}

		#endregion

	}

	#endregion

	#region Class: MetaDataParentInfo

	/// <summary>
	/// Class indicate meta datas of <see cref="ISynchronizationController"/>.
	/// </summary>
	public class MetaDataInfo
	{

		#region Field: Protected

		protected string ForeignColumnName;

		#endregion

		#region Properties: Public

		/// <summary>
		/// Parent schema name.
		/// </summary>
		public string ParentSchemaName { get; protected set; }

		/// <summary>
		/// Parent RemoteId.
		/// </summary>
		public string ParentRemoteId { get; protected set; }

		/// <summary>
		/// Parent Id.
		/// </summary>
		public Guid ParentId { get; protected set; }

		/// <summary>
		/// Entity Id.
		/// </summary>
		public Guid EntityId { get; protected set; }

		/// <summary>
		/// Entity schema name.
		/// </summary>
		public string EntitySchemaName { get; protected set; }

		/// <summary>
		/// Metadata action.
		/// </summary>
		public SyncAction SyncAction { get; protected set; }

		/// <summary>
		/// Entity modified on column value.
		/// </summary>
		public DateTime ModifiedOn { get; protected set; }

		/// <summary>
		/// Metadata owner contact identifier.
		/// </summary>
		public Guid UserContactId { get; protected set; }

		/// <summary>
		/// Entity columns values collection.
		/// </summary>
		public ImmutableDictionary<string, object> ColumnValues { get; }

		#endregion

		#region Constructor: Public

		public MetaDataInfo(IDictionary<string, object> parameters) {
			SyncAction = (SyncAction)parameters["SyncAction"];
			EntityId = (Guid)parameters["EntityId"];
			EntitySchemaName = parameters["EntitySchemaName"].ToString();
			ModifiedOn = (DateTime)parameters["ModifiedOn"];
			UserContactId = (Guid)parameters["UserContactId"];
			ColumnValues = (ImmutableDictionary<string, object>)parameters["ColumnValues"];
			SetForeignColumnName();
		}

		#endregion

		#region Methods: Private

		protected virtual void SetForeignColumnName() {
			switch (EntitySchemaName) {
				case "ActivityParticipant":
					ForeignColumnName = "ActivityId";
					break;
				case "Activity":
					ForeignColumnName = "Id";
					break;
				default: break;
			}
		}

		#endregion

		#region Method: Public

		/// <summary>
		/// Set 
		/// </summary>
		/// <param name="userConnection">UserConnection see <see cref="UserConnection"/></param>
		public virtual void SetParentProperties(UserConnection userConnection) {
			ParentId = (Guid)ColumnValues[ForeignColumnName];
			var select = new Select(userConnection)
					.Column("RemoteId")
					.Column("SyncSchemaName")
					.From("SysSyncMetaData").As("SSMD")
					.Where("SSMD", "LocalId").IsEqual(Column.Parameter(ParentId))
						.And("SSMD", "CreatedById").IsEqual(Column.Parameter(UserContactId)) as Select;
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
					if (reader.Read()) {
						ParentRemoteId = reader.GetColumnValue<string>("RemoteId");
						ParentSchemaName = reader.GetColumnValue<string>("SyncSchemaName");
					}
				}
			}
		}

		#endregion

	}

	#endregion
}