namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using System.Text;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Common;
	using Terrasoft.Common.Json;
	using Terrasoft.Nui.ServiceModel.Extensions;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeActivitySyncProvider

	public abstract class ExchangeActivitySyncProvider : ExchangeSyncProvider
	{

		#region Constructors: Protected

		protected ExchangeActivitySyncProvider(Guid storeId, TimeZoneInfo timeZone, string senderEmailAddress)
			: base(storeId, timeZone, senderEmailAddress) {
		}

		#endregion

		#region Properties: Protected

		protected Guid ActivityCategoryId {
			get;
			set;
		}

		protected ActivityExchangeSettings UserSettings {
			get;
			set;
		}

		#endregion

		#region Methods: Private

		private IEnumerable<LocalItem> GetNotSyncedActivities(SyncContext context, SyncItemSchema primarySchema) {
			string primarySchemaName = primarySchema.PrimaryEntityConfig.SchemaName;
			UserConnection userConnection = context.UserConnection;
			var esq = new EntitySchemaQuery(userConnection.EntitySchemaManager, primarySchemaName);
			esq.PrimaryQueryColumn.IsAlwaysSelect = true;
			AddActivityQueryColumns(esq);
			AddExportFiltersBySettings(context, esq);
			Select select = esq.GetSelectQuery(userConnection);
			AddExportFilters(userConnection, select);
			EntityCollection entities = esq.GetEntityCollection(userConnection);
			if (entities.Count > 0) {
				select.BuildParametersAsValue = true;
				context?.LogInfo(SyncAction.None, SyncDirection.Download, "loaded {0} activities from bpm, Select: {1}", entities.Count, select.GetSqlText());
			}
			foreach (Entity entity in entities) {
				var localItem = new LocalItem(primarySchema);
				localItem.Entities[primarySchemaName].Add(new SyncEntity(entity, SyncState.New));
				yield return localItem;
			}
		}

		private static void AddDynamicGroupFilters(
			IDictionary<string, Guid> localFolderUIds, UserConnection userConnection,
			EntitySchemaQueryFilterCollection filters) {
			if (!localFolderUIds.Any()) {
				return;
			}
			EntitySchemaManager entitySchemaManager = userConnection.EntitySchemaManager;
			var foldersEsq = new EntitySchemaQuery(entitySchemaManager, "ActivityFolder");
			string searchDataColumnName = foldersEsq.AddColumn("SearchData").Name;
			string[] folderIdsStrArray =
				(from folderId in localFolderUIds.Values select folderId.ToString()).ToArray();
			foldersEsq.Filters.Add(foldersEsq.CreateFilterWithParameters(FilterComparisonType.Equal, false,
				"Id", folderIdsStrArray));
			EntityCollection folderEntities = foldersEsq.GetEntityCollection(userConnection);
			if (!folderEntities.Any()) {
				return;
			}
			EntitySchema entitySchema = entitySchemaManager.GetInstanceByName("Activity");
			Guid schemaUId = entitySchema.UId;
			foreach (Entity folderEntity in folderEntities) {
				byte[] data = folderEntity.GetBytesValue(searchDataColumnName);
				string serializedFilters = Encoding.UTF8.GetString(data, 0, data.Length);
				var dataSourceFilters = Json.Deserialize<Terrasoft.Nui.ServiceModel.DataContract.Filters>(
											serializedFilters);
				IEntitySchemaQueryFilterItem esqFilters = dataSourceFilters.BuildEsqFilter(schemaUId, userConnection);
				if (esqFilters != null) {
					filters.Add(esqFilters);
				}
			}
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Updates last date synchronization.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="context">Synchronization context.</param>
		protected abstract void UpdateLastSyncDate(UserConnection userConnection, SyncContext context);

		#endregion

		#region Methods:Public
		
		/// <summary>
		/// Adds columns to activity query.
		/// </summary>
		/// <param name="esq"><see cref="EntitySchemaQuery"/> in which the columns will be added.</param>
		public virtual void AddActivityQueryColumns(EntitySchemaQuery esq) {
			esq.AddColumn("Title");
			esq.AddColumn("StartDate");
			esq.AddColumn("DueDate");
			esq.AddColumn("Status");
			esq.AddColumn("Priority");
			esq.AddColumn("RemindToOwner");
			esq.AddColumn("RemindToOwnerDate");
			esq.AddColumn("Notes");
			esq.AddColumn("Location");
		}
		
		/// <summary>
		/// Adds filters depending on the activity export settings.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="esq"><see cref="EntitySchemaQuery"/> in which the filters will be added.</param>
		/// <remarks>If the directories are not specified for export and the other export options have value "false"
		/// esq have to return empty collection.</remarks>
		public virtual void AddExportFiltersBySettings(SyncContext context, EntitySchemaQuery esq) {
			UserConnection userConnection = context.UserConnection;
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.NotEqual, false, "Type",
					ActivityConsts.EmailTypeUId));
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "ShowInScheduler",
					ActivityCategoryId == ExchangeConsts.ActivityMeetingCategoryId));
			var importActivityFrom = UserSettings.ImportActivitiesFrom;
			if (importActivityFrom != DateTime.MinValue) {
				esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Greater, "DueDate",
					importActivityFrom));
			} 
			if (UserSettings.ExportActivitiesAll) {
				return;
			}
			if (UserSettings.ExportActivitiesFromScheduler && !UserSettings.ExportActivitiesFromGroups) {
				esq.Filters.Add(
					esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "ShowInScheduler", true));
				return;
			}
			if (!UserSettings.ExportActivitiesFromGroups) {
				return;
			}
			if (!UserSettings.LocalFolderUIds.Any() && !UserSettings.ExportActivitiesFromScheduler) {
				esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, 
					esq.RootSchema.PrimaryColumn.Name, Guid.Empty));
				return;
			}
			var dynamicGroupsFilters = new EntitySchemaQueryFilterCollection(esq, LogicalOperationStrict.Or);
			AddDynamicGroupFilters(UserSettings.LocalFolderUIds, userConnection, dynamicGroupsFilters);
			if (UserSettings.ExportActivitiesFromScheduler) {
				dynamicGroupsFilters.Add
					(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "ShowInScheduler", true));
			}
			esq.Filters.Add(dynamicGroupsFilters);
		}
		
		/// <summary>
		/// Adds additional export filters.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="select"><see cref="Select"/> in which the filters will be added.</param>
		public virtual void AddExportFilters(UserConnection userConnection, Select select) {
			var currentUserId = userConnection.CurrentUser.ContactId;
			select.And().Not().Exists(
					new Select(userConnection)
						.Column("LocalId")
					.From("SysSyncMetaData").As("SSMD")
					.Where("SSMD", "SyncSchemaName").IsEqual(new QueryParameter("Activity"))
					.And("SSMD", "LocalId").IsEqual("Activity", "Id")
					.And()
						.OpenBlock("SSMD", "CreatedById").IsEqual(new QueryParameter(
							userConnection.CurrentUser.ContactId))
							.Or()
								.OpenBlock("SSMD", "RemoteItemName").IsNotEqual(new QueryParameter(
									KnownTypes().First().Name))
									.And("SSMD", "CreatedById").IsNotEqual(new QueryParameter(
										userConnection.CurrentUser.ContactId))
								.CloseBlock()
						.CloseBlock()
				);
			if (ActivityCategoryId == ActivityConsts.AppointmentActivityCategoryId) {
				select.InnerJoin("ActivityParticipant").On("Activity", "Id").IsEqual("ActivityParticipant", "ActivityId")
				.And("ActivityParticipant", "ParticipantId").IsEqual(Column.Parameter(currentUserId));
			} else {
				select.And("OwnerId").IsEqual(Column.Parameter(currentUserId));
			}
		}
		
		/// <summary>
		/// Returns an enumerator metadata synchronization objects,
		/// that have been modified in the local store with the last synchronization date.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="modifiedItemsEsq">Request to the scheme <see cref="SysSyncMetaData"/>.</param>
		/// <remarks>The method is overridden.</remarks>
		/// <returns>The enumerator object synchronization of external data storage.</returns>
		public override IEnumerable<ItemMetadata> GetLocallyModifiedItemsMetadata(SyncContext context,
						EntitySchemaQuery modifiedItemsEsq) {
			modifiedItemsEsq.Filters.Add(modifiedItemsEsq.CreateFilterWithParameters(FilterComparisonType.Equal,
					"CreatedBy", context.UserConnection.CurrentUser.ContactId));
			return base.GetLocallyModifiedItemsMetadata(context, modifiedItemsEsq);
		}
		
		/// <summary>
		/// Gets user settings.
		/// </summary>
		/// <returns>User settings</returns>
		public ActivityExchangeSettings GetUserSettings() {
			return UserSettings;
		}

		/// <summary>
		/// Returns have not yet synchronized items added since the last synchronization.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override IEnumerable<LocalItem> CollectNewItems(SyncContext context) {
			SyncItemSchema primarySchema = SyncItemSchemaCollection.First(schema =>
				schema.PrimaryEntityConfig.Order == 0);
			return GetNotSyncedActivities(context, primarySchema);
		}

		/// <summary>
		/// Updates last synchtonization date.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override void CommitChanges(SyncContext context) {
			UpdateLastSyncDate(context.UserConnection, context);
		}

		#endregion
	}

	 #endregion

	#region Class: ActivityExchangeSettings

	/// <summary>
	/// Provides user settings to synchronize data provider.
	/// </summary>
	public abstract class ActivityExchangeSettings : ExchangeSettings
	{
		#region Constructors: Public

		public ActivityExchangeSettings(UserConnection userConnection, string senderEmailAddress)
			: base(userConnection, senderEmailAddress) {
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// Import activities from.
		/// </summary>
		private DateTime _importActivitiesFrom;
		public DateTime ImportActivitiesFrom {
			get {
				if (ActivitySyncPeriod.IsNotEmpty()) {
					var type = LoadFromDateType.GetInstance(_userConnection);
					var date = type.GetLoadFromDate(ActivitySyncPeriod);
					return date.GetUserDateTime(_userConnection).Date;
				}
				return _importActivitiesFrom;
			}
			protected set => _importActivitiesFrom = value;
		}

		/// <summary>
		/// <see cref="MailSyncPeriod"/> instance unique identifier.
		/// </summary>
		public Guid ActivitySyncPeriod { get; set; }

		/// <summary>
		/// Flag, indiacates if need to import activities from Exchange.
		/// </summary>
		public bool ImportActivities {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to import activities from Exchange.
		/// </summary>
		public bool ImportActivitiesAll {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to import activities from Exchange folders.
		/// </summary>
		public bool ImportActivitiesFromFolders {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export activities.
		/// </summary>
		public bool ExportActivities {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export all activities.
		/// </summary>
		public bool ExportActivitiesAll {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export all activities.
		/// </summary>
		public bool ExportActivitiesSelected {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export activities with type "Task" and category "Appointment".
		/// </summary>
		public bool ExportAppointments {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export activities with type "Task" and category "Task".
		/// </summary>
		public bool ExportTasks {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export activities from the schedule.
		/// </summary>
		/// <remarks>
		/// Exports activitties with type "Task" and category "Task" or "Task".
		/// </remarks>
		public bool ExportActivitiesFromScheduler {
			get;
			protected set;
		}

		/// <summary>
		/// Flag, indiacates if need to export activities from the groups.
		/// </summary>
		public bool ExportActivitiesFromGroups {
			get;
			protected set;
		}

		#endregion

		#region Methods: Protected

		protected static IDictionary<string, Guid> GetFoldersInfo(string serializedFolders) {
			var result = new Dictionary<string, Guid>();
			var foldersArray = Json.Deserialize(serializedFolders) as JArray;
			if (foldersArray == null) {
				return result;
			}
			foreach (JToken folder in foldersArray) {
				result.Add(folder.Value<string>("Path"), new Guid(folder.Value<string>("Id")));
			}
			return result;
		}

		#endregion

	}

	#endregion
}