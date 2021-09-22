namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Common;
	using Terrasoft.Sync;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using ItemNotFoundException = Terrasoft.Common.ItemNotFoundException;
	using SysSyncMetaData = Terrasoft.Core.Configuration.SysSyncMetaData;
	using SysSettings = Terrasoft.Core.Configuration.SysSettings;

	#region Class: ExchangeAppointmentSyncProvider

	[DefaultBinding(typeof(BaseExchangeSyncProvider), Name = "ExchangeAppointmentSyncProvider")]
	public class ExchangeAppointmentSyncProvider : ExchangeActivitySyncProvider
	{

		#region Fields: Private

		private static readonly Dictionary<Type, Type> _syncTypeToTypeMap = new Dictionary<Type, Type> {
			{ typeof(ExchangeAppointment), typeof(Microsoft.Exchange.WebServices.Data.Appointment) }
		};

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private UserConnection _userConnection;

		/// <summary>
		/// <see cref="Terrasoft.Configuration.SynchronizationErrorHelper"/> instance.
		/// </summary>
		private readonly SynchronizationErrorHelper _syncErrorHelper;

		#endregion

		#region Fields: Protected

		/// <summary>
		/// Exchange folders for synchronization list.
		/// </summary>
		protected List<Exchange.Folder> Folders;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// It represents a class for performing synchronization operations on objects.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		public ExchangeAppointmentSyncProvider(UserConnection userConnection, string senderEmailAddress)
			: base(ExchangeConsts.AppointmentStoreId, userConnection.CurrentUser.TimeZone, senderEmailAddress) {
			UserSettings = new AppointmentExchangeSettings(userConnection, senderEmailAddress);
			SetVersion(userConnection);
			ActivityCategoryId = ExchangeConsts.ActivityMeetingCategoryId;
			_userConnection = userConnection;
			_syncErrorHelper = SynchronizationErrorHelper.GetInstance(userConnection);
		}

		/// <summary>
		/// Initializes new instance <see cref="ExchangeAppointmentSyncProvider" />.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="timeZoneInfo">Current timezone.</param>
		/// <param name="senderEmailAddress">The mailing address of the synchronization.</param>
		public ExchangeAppointmentSyncProvider(UserConnection userConnection, string senderEmailAddress, TimeZoneInfo timeZoneInfo)
			: base(ExchangeConsts.AppointmentStoreId, timeZoneInfo, senderEmailAddress) {
			_userConnection = userConnection;
		}

		#endregion

		#region Methods:Private

		/// <summary>
		/// Gets filters for Exchange data query.
		/// </summary>
		/// <returns>Filter instance.</returns>
		private Exchange.SearchFilter GetItemsSearchFilters(SyncContext context = null) {
			Exchange.SearchFilter filter = null;
			DateTime lastSyncDateUtc = TimeZoneInfo.ConvertTimeToUtc(UserSettings.LastSyncDate, TimeZone);
			DateTime importActivitiesFromDate = UserSettings.ImportActivitiesFrom;
			context?.LogInfo(SyncAction.None, SyncDirection.Download, "lastSyncDateUtc = {0}, importActivitiesFromDate = {1}",
				lastSyncDateUtc, importActivitiesFromDate);
			if (lastSyncDateUtc != DateTime.MinValue && UserSettings.ImportActivitiesFrom != DateTime.MinValue) {
				var lastSyncDateUtcFilter = new Exchange.SearchFilter.IsGreaterThan(
					Exchange.ItemSchema.LastModifiedTime, lastSyncDateUtc.ToLocalTime());
				var filterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.Or);
				var customPropSetFilter = new Exchange.SearchFilter.Exists(ExchangeUtilityImpl.LocalIdProperty);
				var notCustomPropSetFilter = new Exchange.SearchFilter.Not(customPropSetFilter);
				filterCollection.AddRange(new List<Exchange.SearchFilter> {
					lastSyncDateUtcFilter,
					notCustomPropSetFilter
				});
				if (context != null && GetExchangeRecurringAppointmentsSupported(context.UserConnection)) {
					return filterCollection;
				}
				var allFilterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.And);
				var importActivitiesFromDateFilter = new Exchange.SearchFilter.IsGreaterThanOrEqualTo(
					Exchange.AppointmentSchema.Start, importActivitiesFromDate);
				allFilterCollection.AddRange(new List<Exchange.SearchFilter> {
					importActivitiesFromDateFilter,
					filterCollection
				});
				filter = allFilterCollection;
			} else if (lastSyncDateUtc != DateTime.MinValue && UserSettings.ImportActivitiesFrom == DateTime.MinValue) {
				filter = new Exchange.SearchFilter.IsGreaterThan(
					Exchange.ItemSchema.LastModifiedTime, lastSyncDateUtc.ToLocalTime());
			} else if (lastSyncDateUtc == DateTime.MinValue && UserSettings.ImportActivitiesFrom != DateTime.MinValue) {
				filter = new Exchange.SearchFilter.IsGreaterThanOrEqualTo(
					Exchange.AppointmentSchema.Start, importActivitiesFromDate);
			}
			return filter;
		}

		/// <summary>
		/// Sets current synchronization version.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <remarks>In order to detect appointment delete in exchange calendar, we need to check all
		/// synced appointments for some period. Default period for this operation is seven days,
		/// also it can be changed using DeleteAppointmentsSyncDays setting.</remarks>
		private void SetVersion(UserConnection userConnection) {
			Version = UserSettings.LastSyncDate < UserSettings.ImportActivitiesFrom
				? UserSettings.ImportActivitiesFrom
				: UserSettings.LastSyncDate;
			if (userConnection.GetIsFeatureEnabled("SyncDeletedAppointments")) {
				var deleteAppointmentsSyncDays = SysSettings.GetValue<int>(userConnection, "DeleteAppointmentsSyncDays", 7);
				try {
					Version = Version.AddDays(-1 * deleteAppointmentsSyncDays);
				} catch (ArgumentOutOfRangeException) { }
			}
		}

		/// <summary>
		/// Loads exchange appointment from exchange, filtered by related <paramref name="activityMetadata"/> to
		/// activity unique identifier.
		/// </summary>
		/// <param name="activityMetadata">Activity instance synchronization metadata.</param>
		/// <returns>If appointment with activity id found, returns <see cref="Exchange.Appointment"/> instance.
		/// Otherwise returns null.</returns>
		private Exchange.Appointment GetAppointmentByLocalIdProperty(SysSyncMetaData activityMetadata) {
			var localId = activityMetadata.LocalId;
			var filters = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.And);
			var customPropSetFilter = new Exchange.SearchFilter.IsEqualTo(ExchangeUtilityImpl.LocalIdProperty, localId.ToString());
			filters.Add(customPropSetFilter);
			foreach (var noteFolder in Folders) {
				Exchange.PropertySet idOnlyPropertySet = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
				var itemView = new Exchange.ItemView(1) {
					PropertySet = idOnlyPropertySet
				};
				IEnumerable<Exchange.Item> itemCollection = GetCalendarItemsByFilter(noteFolder, filters, itemView);
				if (itemCollection == null) {
					continue;
				}
				foreach (Exchange.Item item in itemCollection) {
					Exchange.Appointment appointment = SafeBindItem<Exchange.Appointment>(Service, item.Id);
					if (appointment != null) {
						return appointment;
					}
				}
			}
			return null;
		}

		/// <summary>
		/// Creates new <see cref="ExchangeAppointment"/> instance.
		/// </summary>
		/// <param name="schema"><see cref="SyncItemSchema"/> instance.</param>
		/// <param name="item"><see cref="Exchange.Item"/> instance.</param>
		/// <param name="timezone"><see cref="TimeZoneInfo"/> instance.</param>
		/// <returns><see cref="ExchangeAppointment"/> instance.</returns>
		private ExchangeAppointment CreateExchangeAppointment(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timezone) {
			return CreateExchangeAppointment(schema, item, string.Empty, timezone);
		}

		/// <summary>
		/// Creates new <see cref="ExchangeAppointment"/> instance.
		/// </summary>
		/// <param name="schema"><see cref="SyncItemSchema"/> instance.</param>
		/// <param name="item"><see cref="Exchange.Item"/> instance.</param>
		/// <param name="remoteId">Remote item instance unique identifier.</param>
		/// <param name="timezone"><see cref="TimeZoneInfo"/> instance.</param>
		/// <returns><see cref="ExchangeAppointment"/> instance.</returns>
		private ExchangeAppointment CreateExchangeAppointment(SyncItemSchema schema, Exchange.Item item, string remoteId,
				TimeZoneInfo timezone) {
			var schemaParam = new ConstructorArgument("schema", schema);
			var itemParam = new ConstructorArgument("item", item);
			var timeZoneParam = new ConstructorArgument("timeZoneInfo", timezone);
			var remoteIdParam = new ConstructorArgument("remoteId", remoteId);
			return remoteId.IsEmpty()
				? ClassFactory.Get<ExchangeAppointment>(schemaParam, itemParam, timeZoneParam)
				: ClassFactory.Get<ExchangeAppointment>(schemaParam, itemParam, remoteIdParam, timeZoneParam);
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
							.Set("AppointmentLastSyncDate", Column.Parameter(context.CurrentSyncStartVersion))
							.Where("MailboxSyncSettingsId")
							.IsEqual(new Select(userConnection).Top(1)
										.Column("Id")
										.From("MailboxSyncSettings")
										.Where("SenderEmailAddress")
										.IsEqual(Column.Parameter(UserSettings.SenderEmailAddress)));
			update.Execute();
		}

		/// <summary>
		/// Returns activity hash stored in activity synchronization metadata.
		/// </summary>
		/// <param name="itemMetadata"><see cref="ItemMetadata"/> instance.</param>
		/// <returns>Stored activity hash.</returns>
		protected virtual string GetActivityHashFromMetadata(ItemMetadata itemMetadata) {
			string result = string.Empty;
			if (itemMetadata != null) {
				SysSyncMetaData activityMetaData = itemMetadata.FirstOrDefault(item => item.SyncSchemaName == "Activity");
				if (activityMetaData != null) {
					string extraParameters = activityMetaData.ExtraParameters;
					if (!extraParameters.IsNullOrEmpty()) {
						JObject activityExtendProperty = JObject.Parse(activityMetaData.ExtraParameters);
						result = (string)activityExtendProperty["ActivityHash"];
					}
				}
			}
			return result;
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.GetCurrentActionIgnored"/>
		/// </summary>
		protected override bool GetCurrentActionIgnored(SyncAction action) {
			if (_userConnection.GetIsFeatureEnabled("SyncDeletedAppointments")) {
				return false;
			}
			return base.GetCurrentActionIgnored(action);
		}

		/// <summary>
		/// Checks is activities export disabled.
		/// </summary>
		/// <returns><c>True</c> if all export optiond disabled. Otherwise returns <c>false</c>.</returns>
		protected virtual bool GetExportDisabled() {
			return !UserSettings.ExportActivities || !UserSettings.ExportActivitiesAll &&
				!UserSettings.ExportActivitiesFromGroups && !UserSettings.ExportActivitiesFromScheduler &&
				!UserSettings.ExportAppointments;
		}

		/// <summary>
		/// Fills selected for synchronization folreds list.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		protected void FillFoldersList(SyncContext context) {
			Folders = new List<Exchange.Folder>();
			if (UserSettings.ImportActivitiesAll) {
				Exchange.Folder rootFolder = Exchange.Folder.Bind(Service, Exchange.WellKnownFolderName.MsgFolderRoot);
				Folders.GetAllFoldersByClass(rootFolder, ExchangeConsts.AppointmentFolderClassName, null);
			} else {
				Folders = SafeBindFolders(Service, UserSettings.RemoteFolderUIds.Keys, context);
			}
		}

		/// <summary>
		/// Checks is exchange recurring appointments support enabled.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns><c>True</c> if exchange recurring appointments support enabled.Otherwise returns <c>false</c>.</returns>
		protected bool GetExchangeRecurringAppointmentsSupported(UserConnection userConnection) {
			return userConnection.GetIsFeatureEnabled("ExchangeRecurringAppointments");
		}

		/// <summary>
		/// Checks is recurring appointments need to be synced.
		/// Returns <c>true</c> if previous synchronization was in different day.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c> if previous synchronization was in different day. Otherwise returns <c>false</c>.</returns>
		protected bool NeedGetRecurringAppointments(SyncContext context) {
			if (!GetExchangeRecurringAppointmentsSupported(context.UserConnection)) {
				return false;
			}
			var lastSyncDate = UserSettings.LastSyncDate;
			var userNow = DateTime.UtcNow.GetUserDateTime(context.UserConnection);
			return lastSyncDate.Date < userNow.Date;
		}

		/// <summary>
		/// Returns all appointments for period from <paramref name="folder"/>. Recurring apointments included.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="folder"><see cref=" Exchange.Folder"/> instance.</param>
		/// <returns><see cref="Exchange.Item"/> collection.</returns>
		protected IEnumerable<Exchange.Item> GetAppointmentsForPeriod(SyncContext context, Exchange.Folder folder) {
			if (GetExchangeRecurringAppointmentsSupported(context.UserConnection)) {
				var type = LoadFromDateType.GetInstance(context.UserConnection);
				var starDate = type.GetLoadFromDate(UserSettings.ActivitySyncPeriod);
				var endDate = type.GetLoadFromDate(UserSettings.ActivitySyncPeriod, true);
				context?.LogInfo(SyncAction.None, SyncDirection.Download, "GetAppointmentsForPeriod with: starDate {0}, endDate {1}.",
						starDate, endDate);
				var currentStartDate = starDate;
				while (currentStartDate < endDate) {
					var calendarItemView = new Exchange.CalendarView(currentStartDate, currentStartDate.AddDays(1), 150);
					context?.LogInfo(SyncAction.None, SyncDirection.Download, "GetAppointmentsForPeriod with: currentStartDate {0}.",
						currentStartDate);
					Exchange.FindItemsResults<Exchange.Appointment> calendarItemCollection;
					calendarItemCollection = Service.FindAppointments(folder.Id, calendarItemView);
					context?.LogInfo(SyncAction.None, SyncDirection.Download, 
							"GetAppointmentsForPeriod - Service found {0} items.",
							calendarItemCollection.Count());
					foreach (var item in calendarItemCollection) {
						yield return item;
					}
					currentStartDate = currentStartDate.AddDays(1);
				}
				context?.LogInfo(SyncAction.None, SyncDirection.Download, "GetAppointmentsForPeriod - End method.");
			}
		}

		/// <summary>
		/// Returns all items by <paramref name="filter"/> from <paramref name="folder"/>. Recurring apointments not included.
		/// </summary>
		/// <param name="folder"><see cref=" Exchange.Folder"/> instance.</param>
		/// <param name="filter"><see cref=" Exchange.SearchFilter"/> instance.</param>
		/// <returns><see cref="Exchange.Item"/> collection.</returns>
		protected IEnumerable<Exchange.Item> GetCalendarItems(Exchange.Folder folder, Exchange.SearchFilter filter) {
			var itemView = new Exchange.ItemView(PageItemCount);
			Exchange.FindItemsResults<Exchange.Item> itemCollection;
			do {
				itemCollection = folder.ReadItems(filter, itemView);
				foreach (var item in itemCollection) {
					yield return item;
				}
			} while (itemCollection.MoreAvailable);
		}

		protected bool CheckItemInSyncPeriod(SyncContext context, Exchange.Appointment appointment) {
			return !GetExchangeRecurringAppointmentsSupported(context.UserConnection) ||
				appointment.SafeGetValue<DateTime>(Exchange.AppointmentSchema.Start) > UserSettings.ImportActivitiesFrom;
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
		/// Returns have not yet synchronized items added since the last synchronization.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override IEnumerable<LocalItem> CollectNewItems(SyncContext context) {
			if (GetExportDisabled()) {
				return new List<LocalItem>();
			}
			return base.CollectNewItems(context);
		}

		/// <summary>
		/// Returns new instance <see cref="ExchangeAppointment"/>.
		/// </summary>
		/// <param name="schema">The instance of schema<see cref="SyncItemSchema"/>, describing created 
		/// synchronization item in external storage <see cref="ExchangeAppointment"/>.</param>
		/// <returns>New item <see cref="ExchangeAppointment"/>.</returns>
		public override IRemoteItem CreateNewSyncItem(SyncItemSchema schema) {
			var appointment = CreateExchangeAppointment(schema, new Exchange.Appointment(Service), TimeZone);
			appointment.Action = SyncAction.Create;
			return appointment;
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.EnumerateChanges"/>
		/// </summary>
		public override IEnumerable<IRemoteItem> EnumerateChanges(SyncContext context) {
			base.EnumerateChanges(context);
			FillFoldersList(context);
			var result = new List<IRemoteItem>();
			if (!UserSettings.ImportActivities) {
				return result;
			}
			Exchange.SearchFilter itemsFilter = GetItemsSearchFilters(context);
			SyncItemSchema schema = FindSchemaBySyncValueName(typeof(ExchangeAppointment).Name);
			var needGetRecurringAppointments = NeedGetRecurringAppointments(context);
			Exchange.PropertySet properties = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
			properties.AddRange(new[] {
				Exchange.ItemSchema.Subject,
				Exchange.ItemSchema.LastModifiedTime,
				Exchange.AppointmentSchema.Recurrence,
				Exchange.AppointmentSchema.IsRecurring,
				Exchange.AppointmentSchema.AppointmentType,
				Exchange.AppointmentSchema.ICalUid,
				Exchange.AppointmentSchema.ICalRecurrenceId,
				Exchange.AppointmentSchema.Start
			});
			foreach (Exchange.Folder folder in Folders) {
				bool hasRecurence = false;
				if (!needGetRecurringAppointments) {
					foreach (Exchange.Item item in GetCalendarItems(folder, itemsFilter)) {
						Exchange.Appointment appointment = ExchangeUtility.SafeBindItem<Exchange.Appointment>(Service, item.Id, properties);
						if (appointment != null) {
							var recurentAppointment = GetExchangeRecurringAppointmentsSupported(context.UserConnection) && appointment.Recurrence != null;
							var isRecurringMaster = appointment.AppointmentType == Exchange.AppointmentType.RecurringMaster;
							if ((!recurentAppointment || isRecurringMaster) && CheckItemInSyncPeriod(context, appointment)) {
								context?.LogInfo(SyncAction.None, SyncDirection.Download, "Adding single or master appoitment {0}", appointment.Subject);
								var remoteItem = CreateExchangeAppointment(schema, appointment, TimeZone);
								context?.LogInfo(SyncAction.None, SyncDirection.Download, "Created ExchangeAppointment with Id {0}", remoteItem.Id);
								if (isRecurringMaster) {
									context?.LogInfo(SyncAction.None, SyncDirection.Download, "Adding master appoitment {0}", appointment.Subject);
									remoteItem.Action = SyncAction.CreateRecurringMaster;
								}
								result.Add(remoteItem);
							}
							hasRecurence = hasRecurence || recurentAppointment;
						}
					}
				}
				if (hasRecurence || needGetRecurringAppointments) {
					foreach (Exchange.Item item in GetAppointmentsForPeriod(context, folder)) {
						context?.LogInfo(SyncAction.None, SyncDirection.Download, "Input item - Subject {0}, UniqueId - {1}", item.Subject, item.Id.UniqueId);
						Exchange.Appointment appointment = ExchangeUtility.SafeBindItem<Exchange.Appointment>(Service, item.Id, properties);
						context?.LogInfo(SyncAction.None, SyncDirection.Download, "Adding recurence appoitment {0}", appointment.Subject);
						if (appointment != null) {
							var remoteItem = CreateExchangeAppointment(schema, appointment, TimeZone);
							context?.LogInfo(SyncAction.None, SyncDirection.Download, "Created recurence ExchangeAppointment with Id {0}", remoteItem.Id);
							result.Add(remoteItem);
						}
					}
				}
			}
			context?.LogInfo(SyncAction.None, SyncDirection.Download, "loaded {0} items from Exchange", result.Count);
			return result;
		}

		public override IEnumerable<Type> KnownTypes() {
			return _syncTypeToTypeMap.Keys;
		}

		/// <summary>Gets remote Id from extend property.</summary>
		/// <param name="syncItemMetadata">Synchronization metadata.</param>
		/// <returns>Unique identifier item in remote storage.</returns>
		public string GetRemoteId(SysSyncMetaData syncItemMetadata) {
			if (!(syncItemMetadata.ExtraParameters).IsNullOrEmpty()) {
				JObject activityExtendProperty = JObject.Parse(syncItemMetadata.ExtraParameters);
				return (string)activityExtendProperty["RemoteId"];
			} else {
				throw new ItemNotFoundException("Extra parameters is empty");
			}
		}

		/// <summary>
		/// Returns the external data repository synchronization object by a unique foreign key from synchronization metadata.
		/// </summary>
		/// <param name="schema">Schema.</param>
		/// <param name="itemMetadata">Synchronization metadata.</param>
		/// <returns>Created instance of <see cref="IRemoteItem"/>.</returns>
		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, ItemMetadata itemMetadata) {
			SysSyncMetaData syncItemMetadata = itemMetadata.FirstOrDefault(item => item.SyncSchemaName == "Activity");
			if (syncItemMetadata == null) {
				var notChangedAppointment = CreateExchangeAppointment(schema, null, string.Empty, TimeZone);
				notChangedAppointment.State = SyncState.None;
				return notChangedAppointment;
			}
			string itemId = GetRemoteId(syncItemMetadata);
			ExchangeBase remoteItem;
			var appointment = SafeBindItem<Exchange.Appointment>(Service, new Exchange.ItemId(itemId));
			if (appointment == null && _userConnection.GetIsFeatureEnabled("SyncDeletedAppointments")) {
				appointment = GetAppointmentByLocalIdProperty(syncItemMetadata);
			}
			if (appointment != null) {
				remoteItem = CreateExchangeAppointment(schema, appointment, TimeZone);
				remoteItem.Action = SyncAction.Update;
			} else {
				string itemRemoteId = syncItemMetadata.RemoteId;
				appointment = new Exchange.Appointment(Service);
				remoteItem = CreateExchangeAppointment(schema, appointment, itemRemoteId, TimeZone);
				remoteItem.State = SyncState.Deleted;
				remoteItem.Action = SyncAction.Delete;
			}
			return remoteItem;
		}

		/// <summary>In case of items with same remote ids, some new exchange items resolved as apply to remote.
		/// In case there is no activity metadata, conflict resolved as apply to local.</summary>
		/// <param name="syncItem">Sync entity instance.</param>
		/// <param name="itemMetaData">Metadata instance.</param>
		/// <param name="localStoreId">Local storage id.</param>
		/// <returns>Sync conflict resoluton for sync item.</returns>
		public override SyncConflictResolution ResolveConflict(IRemoteItem syncItem,
				ItemMetadata itemMetaData, Guid localStoreId) {
			if (syncItem.Action == SyncAction.CreateRecurringMaster) {
				return SyncConflictResolution.ApplyToLocal;
			}
			bool isActivityExists = itemMetaData.Any(item => item.SyncSchemaName == "Activity");
			if (!isActivityExists) {
				return SyncConflictResolution.ApplyToLocal;
			}
			if (_userConnection.GetIsFeatureEnabled("ResolveConflictsByContent")) {
				var appointment = syncItem as ExchangeAppointment;
				if (appointment.Action == SyncAction.Delete) {
					return SyncConflictResolution.ApplyToRemote;
				}
				Exchange.Appointment remoteItem = appointment.Item as Exchange.Appointment;
				appointment.LoadItemProperties(remoteItem);
				string oldActivityHash = GetActivityHashFromMetadata(itemMetaData);
				if (!appointment.IsAppointmentChanged(remoteItem, oldActivityHash, _userConnection)) {
					return SyncConflictResolution.ApplyToRemote;
				}
			}
			SyncConflictResolution result = base.ResolveConflict(syncItem, itemMetaData, localStoreId);
			return result;
		}

		/// <summary>
		/// Updates last synchtonization date.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		public override void CommitChanges(SyncContext context) {
			base.CommitChanges(context);
			_syncErrorHelper.CleanUpSynchronizationError(SenderEmailAddress);
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.OnLocalItemAppliedInRemoteStore"/>
		/// </summary>
		public override void OnLocalItemAppliedInRemoteStore(SyncContext context, IRemoteItem remoteItem,
				LocalItem localItem) {
			if (!_userConnection.GetIsFeatureEnabled("SyncDeletedAppointments")) {
				return;
			}
			if (remoteItem.Action == SyncAction.Delete && remoteItem.State == SyncState.Deleted) {
				EntitySchema schema = context.UserConnection.EntitySchemaManager.GetInstanceByName("Activity");
				IEnumerable<Guid> entityIds = localItem.Entities["Activity"]
					.Where(se => se.State != SyncState.Deleted).Select(se => se.EntityId).ToList();
				foreach (Guid activityId in entityIds) {
					Entity activity = schema.CreateEntity(context.UserConnection);
					if (activity.FetchFromDB(activityId, false)) {
						localItem.AddOrReplace("Activity", new SyncEntity(activity, SyncState.Modified) {
							Action = SyncAction.Delete
						});
					}
				}
				context.LocalProvider.ApplyChanges(context, localItem);
			}
		}

		/// <summary>
		/// Gets Items from <see cref="folder"/> by <see cref="filterCollection"/>
		/// </summary>
		/// <param name="folder">Exchange folder.</param>
		/// <param name="filterCollection">Filter collection.</param>
		/// <param name="itemView">Represents the view settings in a folder search operation.</param>
		/// <returns>Search result collection.</returns>
		public virtual IEnumerable<Exchange.Item> GetCalendarItemsByFilter(Exchange.Folder folder,
			Exchange.SearchFilter.SearchFilterCollection filterCollection, Exchange.ItemView itemView) {
			return GetFolderItemsByFilter(folder, filterCollection, itemView);
		}

		#endregion
	}

	#endregion

	#region Class: AppointmentExchangeSettings

	/// <summary>
	/// Provides user settings to synchronize data provider.
	/// </summary>
	public class AppointmentExchangeSettings : ActivityExchangeSettings
	{
		#region Constructors: Public

		public AppointmentExchangeSettings(UserConnection userConnection, string senderEmailAddress)
			: base(userConnection, senderEmailAddress) {
		}

		#endregion

		#region Methods: Private

		private void InitSyncConfigData(UserConnection userConnection) {
			var select = new Select(userConnection).Top(1)
							.Column("ss", "AppointmentLastSyncDate")
							.Column("ss", "ImportActivitiesFrom")
							.Column("ss", "ImportAppointments")
							.Column("ss", "ImportAppointmentsAll")
							.Column("ss", "ImpAppointmentsFromCalendars")
							.Column("ss", "ExportActivities")
							.Column("ss", "ExportActivitiesAll")
							.Column("ss", "ExportActivitiesSelected")
							.Column("ss", "ExportAppointments")
							.Column("ss", "ExportActivitiesFromScheduler")
							.Column("ss", "ExportActivitiesFromGroups")
							.Column("ss", "ImportAppointmentsCalendars")
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
						throw new Exception(string.Format("Exchange ActivitySync settings not found." +
							" SenderEmailAddress = '{0}", SenderEmailAddress));
					}
					LastSyncDate = DateTime.SpecifyKind(reader.GetColumnValue<DateTime>("AppointmentLastSyncDate"),
						DateTimeKind.Utc).GetUserDateTime(userConnection);
					ImportActivitiesFrom = DateTime.SpecifyKind(
						reader.GetColumnValue<DateTime>("ImportActivitiesFrom").ToUniversalTime(),
						DateTimeKind.Utc).GetUserDateTime(userConnection);
					if (userConnection.GetIsFeatureEnabled("ExchangeRecurringAppointments")) {
						ActivitySyncPeriod = reader.GetColumnValue<Guid>("ActivitySyncPeriodId");
					}
					ImportActivities = reader.GetColumnValue<bool>("ImportAppointments");
					ImportActivitiesAll = reader.GetColumnValue<bool>("ImportAppointmentsAll");
					ImportActivitiesFromFolders = reader.GetColumnValue<bool>("ImpAppointmentsFromCalendars");
					ExportActivities = reader.GetColumnValue<bool>("ExportActivities");
					ExportActivitiesAll = reader.GetColumnValue<bool>("ExportActivitiesAll");
					ExportActivitiesSelected = reader.GetColumnValue<bool>("ExportActivitiesSelected");
					ExportAppointments = reader.GetColumnValue<bool>("ExportAppointments");
					ExportActivitiesFromScheduler = reader.GetColumnValue<bool>("ExportActivitiesFromScheduler");
					ExportActivitiesFromGroups = reader.GetColumnValue<bool>("ExportActivitiesFromGroups");
					RemoteFolderUIds = GetFoldersInfo(reader.GetColumnValue<string>("ImportAppointmentsCalendars"));
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
