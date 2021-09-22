namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Common;
	using Exchange = Microsoft.Exchange.WebServices.Data;
    using Terrasoft.Core.Factories;

    #region Class: ExchangeTaskSyncProvider

    /// <summary>
    /// Class for synchronization with Exchange tasks.
    /// </summary>
    [DefaultBinding(typeof(BaseExchangeSyncProvider), Name = "ExchangeTaskSyncProvider")]
	public class ExchangeTaskSyncProvider : ExchangeActivitySyncProvider
	{

		#region Fields: Private

		private static readonly Dictionary<Type, Type> _syncTypeToTypeMap = new Dictionary<Type, Type> {
			{ typeof(ExchangeTask), typeof(Microsoft.Exchange.WebServices.Data.Task) }
		};

		/// <summary>
		/// <see cref="Terrasoft.Core.UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		/// <summary>
		/// <see cref="Terrasoft.Configuration.SynchronizationErrorHelper"/> instance.
		/// </summary>
		private readonly SynchronizationErrorHelper _syncErrorHelper;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Returns inited <see cref="ExchangeTaskSyncProvider"/> instance.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="settings"><see cref="ActivityExchangeSettings"/> instance.</param>
		public ExchangeTaskSyncProvider(UserConnection userConnection, string senderEmailAddress, ActivityExchangeSettings settings = null)
				: base(ExchangeConsts.ExchangeTaskStoreId, userConnection.CurrentUser.TimeZone, senderEmailAddress) {
			_userConnection = userConnection;
			UserSettings = settings ?? new TaskExchangeSettings(userConnection, senderEmailAddress);
			if (UserSettings.LastSyncDate < UserSettings.ImportActivitiesFrom) {
				Version = UserSettings.ImportActivitiesFrom;
			} else {
				Version = UserSettings.LastSyncDate;
			}
			ActivityCategoryId = ExchangeConsts.ActivityTaskCategoryId;
			_syncErrorHelper = SynchronizationErrorHelper.GetInstance(userConnection);
		}

		#endregion
			
		#region Methods:Private

		/// <summary>
		/// Gets filters for Exchange data query.
		/// </summary>
		/// <returns>Filter instance.</returns>
		private Exchange.SearchFilter.SearchFilterCollection GetItemsSearchFilters() {
			var itemsFilter = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.And) {
				new Exchange.SearchFilter.IsEqualTo(Exchange.ItemSchema.ItemClass, ExchangeConsts.TaskItemClassName)
			};
			DateTime lastSyncDateUtc = TimeZoneInfo.ConvertTimeToUtc(UserSettings.LastSyncDate, TimeZone);
			DateTime importActivitiesFromDate = UserSettings.ImportActivitiesFrom;
			if (UserSettings.ImportActivitiesFrom != DateTime.MinValue) {
				itemsFilter.Add(new Exchange.SearchFilter.IsGreaterThanOrEqualTo(
					Exchange.TaskSchema.LastModifiedTime, importActivitiesFromDate));
			}			
			if (UserSettings.LastSyncDate != DateTime.MinValue) {
				var lastSyncDateUtcFilter = new Exchange.SearchFilter.IsGreaterThan(
					Exchange.ItemSchema.LastModifiedTime, lastSyncDateUtc.ToLocalTime());
				var filterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.Or);
				var customPropSetFilter = new Exchange.SearchFilter.Exists(ExchangeUtilityImpl.LocalIdProperty);
				var notCustomPropSetFilter = new Exchange.SearchFilter.Not(customPropSetFilter);
				filterCollection.AddRange(new List<Exchange.SearchFilter> {
					lastSyncDateUtcFilter,
					notCustomPropSetFilter
				});
				itemsFilter.Add(filterCollection);
			}
			return itemsFilter;
		}
			
		#endregion

		#region Methods: Protected

		/// <summary>
		/// Updates last synchronization date.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="context">Synchronization context.</param>
		protected override void UpdateLastSyncDate(UserConnection userConnection, SyncContext context) {
			var update = new Update(userConnection, "ActivitySyncSettings")
					.Set("TaskLastSyncDate", Column.Parameter(context.CurrentSyncStartVersion))
					.Where("MailboxSyncSettingsId")
					.IsEqual(new Select(userConnection).Top(1)
							.Column("Id")
							.From("MailboxSyncSettings")
							.Where("SenderEmailAddress")
							.IsEqual(Column.Parameter(UserSettings.SenderEmailAddress)));
			update.Execute();
		}

		#endregion

		#region Methods:Public

		/// <summary>
		/// <see cref="RemoteProvider.NeedMetaDataActualization"/>
		/// </summary>
		public override bool NeedMetaDataActualization() {
			return !_userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata");
		}

		/// <summary>
		/// Returns items that haven't been synchronized since the last synchronization.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override IEnumerable<LocalItem> CollectNewItems(SyncContext context) {
			if (!UserSettings.ExportActivities || (!UserSettings.ExportActivitiesAll
					&& !UserSettings.ExportActivitiesFromGroups
					&& !UserSettings.ExportActivitiesFromScheduler && !UserSettings.ExportTasks)) {
				return new List<LocalItem>();
			}
			return base.CollectNewItems(context);
		}

		/// <summary>
		/// Returns new instance <see cref="ExchangeTask"/>.
		/// </summary>
		/// <param name="schema">The instance of schema<see cref="SyncItemSchema"/>, describing created 
		/// synchronization item in external storage <see cref="ExchangeTask"/>.</param>
		/// <returns>New item <see cref="ExchangeTask"/>.</returns>
		public override IRemoteItem CreateNewSyncItem(SyncItemSchema schema) {
			return new ExchangeTask(schema, new Exchange.Task(Service), TimeZone) {
				Action = SyncAction.Create
			};
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.EnumerateChanges"/>
		/// </summary>
		public override IEnumerable<IRemoteItem> EnumerateChanges(SyncContext context) {
			base.EnumerateChanges(context);
			var result = new List<IRemoteItem>();
			if (!UserSettings.ImportActivities) {
				return result;
			}
			var folders = new List<Exchange.Folder>();
			Exchange.FolderId trashFolderId = Exchange.Folder.Bind(
					Service, Exchange.WellKnownFolderName.DeletedItems, Exchange.BasePropertySet.IdOnly).Id;
			if (UserSettings.ImportActivitiesAll) {
				Exchange.Folder rootFolder = Exchange.Folder.Bind(Service, Exchange.WellKnownFolderName.MsgFolderRoot);
				folders.GetAllFoldersByFilter(rootFolder);
				folders.Add(rootFolder);
			} else {
				folders = SafeBindFolders(Service, UserSettings.RemoteFolderUIds.Keys, context);
			}
			Exchange.SearchFilter itemsFilter = GetItemsSearchFilters();
			SyncItemSchema schema = FindSchemaBySyncValueName(typeof(ExchangeTask).Name);
			foreach (Exchange.Folder folder in folders) {
				if (folder.Id.Equals(trashFolderId)) {
					continue;
				}
				var itemView = new Exchange.ItemView(PageItemCount);
				Exchange.FindItemsResults<Exchange.Item> itemCollection;
				do {
					itemCollection = folder.ReadItems(itemsFilter, itemView);
					foreach (Exchange.Item item in itemCollection) {
						Exchange.Task task = ExchangeUtility.SafeBindItem<Exchange.Task>(Service, item.Id);
						if (task != null) {
							var remoteItem = new ExchangeTask(schema, task, TimeZone);
							result.Add(remoteItem);
						}
					}
				} while (itemCollection.MoreAvailable);
			}
			return result;
		}

		public override IEnumerable<Type> KnownTypes() {
			return _syncTypeToTypeMap.Keys;
		}

		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, string id) {
			ExchangeBase remoteItem;
			Exchange.Task task = ExchangeUtility.SafeBindItem<Exchange.Task>(Service, new Exchange.ItemId(id));
			if (task != null) {
				remoteItem = new ExchangeTask(schema, task, TimeZone) {
					Action = SyncAction.Update
				};
			} else {
				task = new Exchange.Task(Service);
				remoteItem = new ExchangeTask(schema, task, id, TimeZone) {
					State = SyncState.Deleted
				};
			}
			return remoteItem;
		}

		/// <summary>
		/// Writes error message to the log, including information about the exception that caused this error.
		/// </summary>
		/// <param name="format">Format string with an informational message.</param>
		/// <param name="exception">Exeption.</param>
		/// <param name="args">Message format.</param>
		public override void LogError(string format, Exception exception, params object[] args) {
			base.LogError(format, exception, args);
			_syncErrorHelper.ProcessSynchronizationError(SenderEmailAddress, exception);
		}

		/// <summary>
		/// Updates last synchtonization date.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override void CommitChanges(SyncContext context) {
			base.CommitChanges(context);
			_syncErrorHelper.CleanUpSynchronizationError(SenderEmailAddress);
		}

		#endregion
	}

	#endregion

	#region Class: TaskExchangeSettings

	/// <summary>
	/// Provides user settings for synchronize data provider.
	/// </summary>
	public class TaskExchangeSettings : ActivityExchangeSettings
	{
		#region Constructors: Public

		public TaskExchangeSettings(UserConnection userConnection, string senderEmailAddress)
			: base(userConnection, senderEmailAddress) {
		}

		#endregion

		#region Methods: Private

		private void InitSyncConfigData(UserConnection userConnection) {
			var select = new Select(userConnection).Top(1)
							.Column("ss", "TaskLastSyncDate")
							.Column("ss", "ImportActivitiesFrom")
							.Column("ss", "ImportTasks")
							.Column("ss", "ImportTasksAll")
							.Column("ss", "ImportTasksFromFolders")
							.Column("ss", "ExportActivities")
							.Column("ss", "ExportActivitiesAll")
							.Column("ss", "ExportActivitiesSelected")
							.Column("ss", "ExportTasks")
							.Column("ss", "ExportActivitiesFromScheduler")
							.Column("ss", "ExportActivitiesFromGroups")
							.Column("ss", "ImportTasksFolders")
							.Column("ss", "ExportActivitiesGroups")
							.Column("ss", "ActivitySyncPeriodId")
							.From("ActivitySyncSettings").As("ss")
							.InnerJoin("MailboxSyncSettings").As("ms")
							.On("ms", "Id").IsEqual("ss", "MailboxSyncSettingsId")
							.Where("ms", "SenderEmailAddress").IsEqual(Column.Parameter(SenderEmailAddress))
							.And("ms", "SysAdminUnitId").IsEqual(Column.Parameter(userConnection.CurrentUser.Id)) as Select;
			using (DBExecutor dbExec = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExec)) {
					if (!reader.Read()) {
						throw new Exception(string.Format("Exchange TaskSync settings not found." +
							" SenderEmailAddress = '{0}", SenderEmailAddress));
					}
					LastSyncDate = DateTime.SpecifyKind(reader.GetColumnValue<DateTime>("TaskLastSyncDate"),
						DateTimeKind.Utc).GetUserDateTime(userConnection);
					ImportActivitiesFrom = DateTime.SpecifyKind(
						reader.GetColumnValue<DateTime>("ImportActivitiesFrom").ToUniversalTime(),
						DateTimeKind.Utc).GetUserDateTime(userConnection);
					if (userConnection.GetIsFeatureEnabled("ExchangeRecurringAppointments")) {
						ActivitySyncPeriod = reader.GetColumnValue<Guid>("ActivitySyncPeriodId");
					}
					ImportActivities = reader.GetColumnValue<bool>("ImportTasks");
					ImportActivitiesAll = reader.GetColumnValue<bool>("ImportTasksAll");
					ImportActivitiesFromFolders = reader.GetColumnValue<bool>("ImportTasksFromFolders");
					ExportActivities = reader.GetColumnValue<bool>("ExportActivities");
					ExportActivitiesAll = reader.GetColumnValue<bool>("ExportActivitiesAll");
					ExportActivitiesSelected = reader.GetColumnValue<bool>("ExportActivitiesSelected");
					ExportTasks = reader.GetColumnValue<bool>("ExportTasks");
					ExportActivitiesFromScheduler = reader.GetColumnValue<bool>("ExportActivitiesFromScheduler");
					ExportActivitiesFromGroups = reader.GetColumnValue<bool>("ExportActivitiesFromGroups");
					RemoteFolderUIds = GetFoldersInfo(reader.GetColumnValue<string>("ImportTasksFolders"));
					LocalFolderUIds = GetFoldersInfo(reader.GetColumnValue<string>("ExportActivitiesGroups"));
				}
			}
		}

		#endregion

		#region Methods: Protected

		protected override void InitializeProperties(UserConnection userConnection) {
			InitSyncConfigData(userConnection);
		}

		#endregion

	}

	#endregion

}
