namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Sync;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Common;
	using Terrasoft.Core.Factories;
	using Exchange = Microsoft.Exchange.WebServices.Data;
    using global::Common.Logging;
	using SysSettings = Terrasoft.Core.Configuration.SysSettings;

	/// <summary>
	/// Class for synchronization with Exchange emails.
	/// </summary>
	[DefaultBinding(typeof(BaseExchangeSyncProvider), Name = "ExchangeEmailSyncProvider")]
	public class ExchangeEmailSyncProvider : ExchangeSyncProvider
	{

		#region Constants: Private

		private const string EmailStoreId = "3294DD3A-BE87-4659-B365-E0AB3D11770C";

		/// <summary>
		/// Session information template.
		/// </summary>
		private const string _sessionInfoTemplate = "[SyncSessionId: {0}, utc now: {1}, mailbox: {2}]  ";

		/// <summary>
		/// LastSyncDate minutes offset for filtration emails in synchronization process.
		/// </summary>
		private readonly int _lastSyncDateMinutesOffset;

		/// <summary>
		/// <see cref="Terrasoft.Configuration.SynchronizationErrorHelper"/> instance.
		/// </summary>
		private readonly SynchronizationErrorHelper _syncErrorHelper;

		#endregion

		#region Constants: Protected

		/// <summary>
		/// The number of items processed in bpm as one package.
		/// </summary>
		/// <remarks>It is still <see cref="ExchangeSyncProvider.PageItemCount"/> * 3.</remarks>
		protected new const int PageElementsCount = 123;

		#endregion

		#region Fields: Private

		private readonly List<Type> _syncListTypes = new List<Type> { typeof(ExchangeEmailMessage) };

		#endregion

		#region Fields: Protected

		protected List<Exchange.Folder> _folderCollection;
		protected Exchange.SearchFilter.SearchFilterCollection _itemsFilterCollection;
		protected Exchange.Folder _currentFolder;
		protected bool _isCurrentFolderProcessed;

		/// <summary>
		/// <see cref="Terrasoft.Core.UserConnection"/> instance.
		/// </summary>
		protected UserConnection _userConnection;

		#endregion

		#region Constructors: Public


		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEmailSyncProvider" /> with passed synchronization settings
		/// with load emails from date parameter.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="loadEmailsFromDate">Load emails from date parameter.</param>
		/// <param name="userSettings"><see cref="EmailExchangeSettings"/> instance.</param>
		public ExchangeEmailSyncProvider(UserConnection userConnection, string senderEmailAddress,
				DateTime loadEmailsFromDate, EmailExchangeSettings userSettings = null)
			: this(userConnection, senderEmailAddress, userSettings) {
			LoadEmailsFromDate = loadEmailsFromDate;
		}

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEmailSyncProvider" /> with passed synchronization settings.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="userSettings"><see cref="EmailExchangeSettings"/> instance.</param>
		public ExchangeEmailSyncProvider(UserConnection userConnection, string senderEmailAddress, EmailExchangeSettings userSettings = null)
			: base(new Guid(EmailStoreId), userConnection.CurrentUser.TimeZone, senderEmailAddress) {
			_userConnection = userConnection;
			UserSettings = userSettings ?? new EmailExchangeSettings(userConnection, senderEmailAddress);
			Version = UserSettings.LastSyncDate;
			_lastSyncDateMinutesOffset = UserSettings.LastSyncDateMinutesOffset;
			InitSyncSessionId();
			_syncErrorHelper = SynchronizationErrorHelper.GetInstance(userConnection);
		}

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEmailSyncProvider" />.
		/// </summary>
		/// <param name="timeZoneInfo">Current timezone.</param>
		/// <param name="senderEmailAddress">The mailing address of the synchronization.</param>
		public ExchangeEmailSyncProvider(string senderEmailAddress, TimeZoneInfo timeZoneInfo)
			: base(new Guid(EmailStoreId), timeZoneInfo, senderEmailAddress) {
			InitSyncSessionId();
		}

		#endregion

		#region Properties: Public

		public EmailExchangeSettings UserSettings {
			get;
			private set;
		}

		/// <summary>
		/// Load emails from date parameter.
		/// </summary>
		public DateTime LoadEmailsFromDate {
			get;
			private set;
		}

		/// <summary>
		/// Synchronization session unique identifier.
		/// </summary>
		private string _synsSessionId;
		public string SynsSessionId {
			get {
				return _synsSessionId;
			}
		}

		/// <summary>
		/// <see cref="RemoteProvider.UseMetadata"/>
		/// </summary>
		public override bool UseMetadata => !_userConnection.GetIsFeatureEnabled("DoNotUseMetadataForEmail");

		/// <summary>
		/// Session information.
		/// </summary>
		public override string SessionInfo {
			get {
				return String.Format(_sessionInfoTemplate, SynsSessionId, DateTime.UtcNow, SenderEmailAddress);
			}
		}
		#endregion

		#region Methods: Private

		/// <summary>
		/// Updates last synchronization date value.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		private void UpdateLastSyncDate(UserConnection userConnection, SyncContext context) {
			if (context.CurrentSyncStartVersion > DateTime.MinValue) {
				var update = new Update(userConnection, "MailboxSyncSettings")
					.Set("LastSyncDate", Column.Parameter(context.CurrentSyncStartVersion))
					.Where("SenderEmailAddress").IsEqual(Column.Parameter(UserSettings.SenderEmailAddress));
				update.Execute();
			}
		}

		/// <summary>
		/// Sets <see cref="SyncSessionId"/> property value.
		/// </summary>
		private void InitSyncSessionId() {
			_synsSessionId = string.Format("ExchangeSyncSession_{0}", Guid.NewGuid());
		}

		/// <summary>
		/// Get <see cref="LastSyncDate"/> converting to <see cref="DateTimeKind.Local"/> value
		/// </summary>
		/// <param name="loadEmailsFromDate">Load emails from date parameter.</param>
		/// <returns>Local <see cref="LastSyncDate"/></returns>
		private DateTime GetLastSyncDate(DateTime loadEmailsFromDate) {
			DateTime lastSyncDateUtc = TimeZoneInfo.ConvertTimeToUtc(loadEmailsFromDate, TimeZone);
			DateTime lastSyncDate = lastSyncDateUtc == DateTime.MinValue
					? lastSyncDateUtc
					: lastSyncDateUtc.AddMinutes(_lastSyncDateMinutesOffset);
			return lastSyncDate.ToLocalTime();
		}

		/// <summary>
		/// Get folder filters.
		/// </summary>
		/// <returns>Folder filters.</returns>
		private Exchange.SearchFilter getFolderFilters() {
			var filterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.Or);
			var filter = new Exchange.SearchFilter.IsEqualTo(Exchange.FolderSchema.FolderClass, ExchangeConsts.NoteFolderClassName);
			var nullfilter = new Exchange.SearchFilter.Exists(Exchange.FolderSchema.FolderClass);
			filterCollection.Add(filter);
			filterCollection.Add(new Exchange.SearchFilter.Not(nullfilter));
			return filterCollection;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// <see cref="ExchangeSyncProvider.FillRemoteFoldersList"/>
		/// </summary>
		protected override void FillRemoteFoldersList(ref List<Exchange.Folder> folders, SyncContext context) {
			if (folders != null) {
				return;
			}
			folders = new List<Exchange.Folder>();
			if (UserSettings.LoadAll) {
				var id = new Exchange.FolderId(Exchange.WellKnownFolderName.MsgFolderRoot, UserSettings.SenderEmailAddress);
				Exchange.Folder rootFolder = Exchange.Folder.Bind(Service, id);
				folders.GetAllFoldersByFilter(rootFolder, getFolderFilters());
			} else {
				folders = SafeBindFolders(Service, UserSettings.RemoteFolderUIds.Keys, context);
			}
			FilterDeprecatedFolders(ref folders);
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.FillItemsFilterCollection"/>
		/// </summary>
		protected override void FillItemsFilterCollection() {
			_itemsFilterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.And);
			var draftFilter = new Exchange.SearchFilter.IsNotEqualTo(Exchange.ItemSchema.IsDraft, true);
			_itemsFilterCollection.Add(draftFilter);
			if (_userConnection.GetIsFeatureEnabled("SetEmailSynchronizedInExchange")) {
				var customPropSetFilter = new Exchange.SearchFilter.Exists(ExchangeUtilityImpl.LocalIdProperty);
				var notCustomPropSetFilter = new Exchange.SearchFilter.Not(customPropSetFilter);
				_itemsFilterCollection.Add(notCustomPropSetFilter);
			}
			DateTime loadEmailsFromDate = LoadEmailsFromDate != DateTime.MinValue
				? LoadEmailsFromDate
				: UserSettings.LastSyncDate;
			LogError($"LoadEmailsFromDate = '{LoadEmailsFromDate}', LastSyncDate = '{UserSettings.LastSyncDate}', result = '{loadEmailsFromDate}'");
			if (loadEmailsFromDate != DateTime.MinValue) {
				var localLastSyncDate = GetLastSyncDate(loadEmailsFromDate);
				var itemsFilter = new Exchange.SearchFilter.IsGreaterThan(Exchange.ItemSchema.LastModifiedTime,
					localLastSyncDate);
				_itemsFilterCollection.Add(itemsFilter);
				LogError($"LoadEmailsFromDate filter adedd, filter date '{localLastSyncDate}'");
				AddLessThanSyncDateFilter(localLastSyncDate);
			}
		}

		/// <summary>
		/// Add less than SyncDate filter to <seealso cref="_itemsFilterCollection"/>
		/// </summary>
		/// <param name="localLastSyncDate">Sync date.</param>
		private void AddLessThanSyncDateFilter(DateTime localLastSyncDate) {
			if (!_userConnection.GetIsFeatureEnabled("IsLessThanSyncDateInterval")) {
				return;
			}
			var lessThanSyncDateInterval = SysSettings.GetValue(_userConnection, "LessThanSyncDateInterval", 1);
			var lessThanSyncDate = localLastSyncDate.AddDays(lessThanSyncDateInterval);
			var ffilter = new Exchange.SearchFilter.IsLessThan(Exchange.ItemSchema.LastModifiedTime,
				lessThanSyncDate);
			LogError($"LessThanSyncDate filter adedd, filter date '{lessThanSyncDate}', interval {lessThanSyncDateInterval}");
			_itemsFilterCollection.Add(ffilter);
		}

		/// <summary>
		/// Gets true if the elements of the collection are available for reading.
		/// </summary>
		/// <param name="itemCollection">List exchange items.</param>
		/// <returns>Flag indicates if the elements of the collection are available for reading. </returns>
		protected virtual bool GetMoreAvailable(Exchange.FindItemsResults<Exchange.Item> itemCollection) {
			return itemCollection.MoreAvailable;
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.GetItemsFromFolders"/>
		/// </summary>
		protected override List<IRemoteItem> GetItemsFromFolders() {
			List<IRemoteItem> result = new List<IRemoteItem>();
			int currentFolderIndex = 0;
			if (_currentFolder != null) {
				currentFolderIndex = _folderCollection.FindIndex(x => GetFolderId(x) == GetFolderId(_currentFolder));
			}
			var numberSkippedFolders = Convert.ToInt32(_isCurrentFolderProcessed) + currentFolderIndex;
			foreach (var noteFolder in _folderCollection.Skip(numberSkippedFolders)) {
				_isCurrentFolderProcessed = false;
				_currentFolder = noteFolder;
				var activityFolderIds = new List<Guid>();
				if (UserSettings.RootFolderId != Guid.Empty) {
					activityFolderIds.Add(UserSettings.RootFolderId);
				}
				var folderId = GetFolderId(noteFolder);
				if (UserSettings.RemoteFolderUIds.ContainsKey(folderId)) {
					activityFolderIds.Add(UserSettings.RemoteFolderUIds[folderId]);
				}
				Exchange.PropertySet idOnlyPropertySet = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
				var itemView = new Exchange.ItemView(PageItemCount) {
					PropertySet = idOnlyPropertySet,
					Offset = PageNumber * PageElementsCount
				};
				Exchange.FindItemsResults<Exchange.Item> itemCollection;
				bool isMaxCountElements = false;
				bool isMoreAvailable = false;
				do {
					itemCollection = GetFolderItemsByFilter(noteFolder, _itemsFilterCollection, itemView);
					result.AddRange(GetEmailsFromCollection(itemCollection, activityFolderIds));
					isMaxCountElements = (result.Count < PageElementsCount);
					isMoreAvailable = GetMoreAvailable(itemCollection);
				} while (isMoreAvailable && isMaxCountElements);
				if (!isMoreAvailable) {
					_isCurrentFolderProcessed = true;
					PageNumber = 0;
				}
				if (!isMaxCountElements) {
					PageNumber++;
					break;
				}
			}
			return result;
		}

		/// <summary>
		/// Sends synchronization session finish message.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		protected virtual void SendSyncSessionFinished(UserConnection userConnection) {
			var userConnectionParam = new ConstructorArgument("userConnection", userConnection);
			EmailMessageHelper helper = ClassFactory.Get<EmailMessageHelper>(userConnectionParam);
			helper.SendSyncSessionFinished(SynsSessionId);
		}

		/// <summary>
		/// Creates upload emai attachments job.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <remarks>External dependency realocation.</remarks>
		protected virtual void CreateUploadAttachmentJob(SyncContext context) {
			ExchangeUtility.CreateUploadAttachmentJob(context.UserConnection, UserSettings.SenderEmailAddress);
		}

		/// <summary>
		/// Uploads emai attachments job. 
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <remarks>External dependency realocation.</remarks>
		protected virtual void UploadAttachmentsData(SyncContext context) {
			ExchangeUtility.UploadAttachmentsData(context.UserConnection, UserSettings.SenderEmailAddress);
		}

		/// <summary>
		/// Unlocks processed in current synchronization entities for synchronization in another synchronization sessions.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		protected void UnlockSyncedEntities(UserConnection userConnection) {
			var helper = ClassFactory.Get<EntitySynchronizerHelper>();
			helper.UnlockEntities(userConnection, "EmailSynchronization");
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="RemoteProvider.GetNumberPageElements"/>
		/// </summary>
		public override int GetNumberPageElements(SyncContext context) {
			return PageElementsCount;
		}

		/// <summary>
		/// <see cref="RemoteProvider.RealeaseItem"/>
		/// </summary>
		public override void RealeaseItem(ref IRemoteItem remoteItem) {
			remoteItem = null;
		}

		/// <summary>
		/// <see cref="RemoteProvider.NeedMetaDataActualization"/>
		/// </summary>
		public override bool NeedMetaDataActualization() {
			return false;
		}

		/// <summary>
		/// <see cref="RemoteProvider.ApplyChanges"/>
		/// </summary>
		/// <remarks>
		/// Email integration only loads messages from Exchange.
		/// </remarks>
		public override void ApplyChanges(SyncContext context, IRemoteItem syncItem) {
		}

		/// <summary>
		/// <see cref="RemoteProvider.CommitChanges"/>
		/// </summary>
		/// <remarks>
		/// Creates upload attachments schedule task.
		/// </remarks>
		public override void CommitChanges(SyncContext context) {
			if (context.UserConnection.GetIsFeatureEnabled("LoadAttachmentsInSameProcess")) {
				UploadAttachmentsData(context);
			} else {
				CreateUploadAttachmentJob(context);
			}
			UpdateLastSyncDate(context.UserConnection, context);
			SendSyncSessionFinished(context.UserConnection);
			_syncErrorHelper.CleanUpSynchronizationError(SenderEmailAddress);
			UnlockSyncedEntities(context.UserConnection);
		}

		/// <summary>
		/// <see cref="RemoteProvider.CollectNewItems"/>
		/// </summary>
		/// <remarks>
		/// Email integration only loads messages from Exchange.
		/// </remarks>
		public override IEnumerable<LocalItem> CollectNewItems(SyncContext context) {
			return new List<LocalItem>();
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.EnumerateChanges"/>
		/// </summary>
		public override IEnumerable<IRemoteItem> EnumerateChanges(SyncContext context) {
			base.EnumerateChanges(context);
			FillRemoteFoldersList(ref _folderCollection, context);
			FillItemsFilterCollection();
			return GetItemsFromFolders();
		}

		public IEnumerable<IRemoteItem> EnumerateChangesOld(SyncContext context) {
			base.EnumerateChanges(context);
			var result = new List<IRemoteItem>();
			var folders = new List<Exchange.Folder>();
			if (UserSettings.LoadAll) {
				var id = new Exchange.FolderId(Exchange.WellKnownFolderName.MsgFolderRoot, UserSettings.SenderEmailAddress);
				Exchange.Folder rootFolder = Exchange.Folder.Bind(Service, id);
				folders.GetAllFoldersByFilter(rootFolder, getFolderFilters());
			} else {
				folders = SafeBindFolders(Service, UserSettings.RemoteFolderUIds.Keys, context);
			}
			var itemsFilterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.And);
			var draftFilter = new Exchange.SearchFilter.IsNotEqualTo(Exchange.ItemSchema.IsDraft, true);
			itemsFilterCollection.Add(draftFilter);
			DateTime loadEmailsFromDate = LoadEmailsFromDate != DateTime.MinValue
				? LoadEmailsFromDate
				: UserSettings.LastSyncDate;
			if (loadEmailsFromDate != DateTime.MinValue) {
				var localLastSyncDate = GetLastSyncDate(loadEmailsFromDate);
				var itemsFilter = new Exchange.SearchFilter.IsGreaterThan(Exchange.ItemSchema.LastModifiedTime,
					localLastSyncDate);
				_itemsFilterCollection.Add(itemsFilter);
			}
			FilterDeprecatedFolders(ref folders);
			foreach (var noteFolder in folders) {
				var activityFolderIds = new List<Guid>();
				if (UserSettings.RootFolderId != Guid.Empty) {
					activityFolderIds.Add(UserSettings.RootFolderId);
				}
				var folderId = GetFolderId(noteFolder);
				if (UserSettings.RemoteFolderUIds.ContainsKey(folderId)) {
					activityFolderIds.Add(UserSettings.RemoteFolderUIds[folderId]);
				}
				Exchange.PropertySet idOnlyPropertySet = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
				var itemView = new Exchange.ItemView(PageItemCount) {
					PropertySet = idOnlyPropertySet
				};
				Exchange.FindItemsResults<Exchange.Item> itemCollection;
				do {
					itemCollection = GetFolderItemsByFilter(noteFolder, itemsFilterCollection, itemView);
					if (itemCollection == null) {
						break;
					}
					result.AddRange(GetEmailsFromCollection(itemCollection, activityFolderIds));
				} while (itemCollection.MoreAvailable);
			}
			return result;
		}

		/// <summary>
		/// Gets emails from <see cref="itemCollection"/>
		/// </summary>
		/// <param name="itemCollection">Finding list items from exchange.</param>
		/// <param name="activityFolderIds">List folders uniqueidentifier.</param>
		/// <returns></returns>
		public virtual IEnumerable<IRemoteItem> GetEmailsFromCollection(Exchange.FindItemsResults<Exchange.Item> itemCollection,
			List<Guid> activityFolderIds) {
			SyncItemSchema schema = FindSchemaBySyncValueName(typeof(ExchangeEmailMessage).Name);
			foreach (Exchange.Item item in itemCollection) {
				if (item is Exchange.EmailMessage) {
					Exchange.PropertySet properties = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
					Exchange.EmailMessage bindEmail = ExchangeUtility.SafeBindItem<Exchange.EmailMessage>(Service, item.Id, properties);
					if (bindEmail != null) {
						var remoteItem = new ExchangeEmailMessage(schema, bindEmail, TimeZone) {
							ActivityFolderIds = activityFolderIds
						};
						yield return remoteItem;
					}
				}
			}
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.LoadSyncItem(SyncItemSchema,string)"/>
		/// </summary>
		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, string id) {
			ExchangeBase remoteItem = null;
			Exchange.EmailMessage fullEmail = ExchangeUtility.SafeBindItem<Exchange.EmailMessage>(Service, new Exchange.ItemId(id));
			if (fullEmail != null) {
				remoteItem = new ExchangeEmailMessage(schema, fullEmail, TimeZone);
			} else {
				fullEmail = new Exchange.EmailMessage(Service);
				remoteItem = new ExchangeEmailMessage(schema, fullEmail, id, TimeZone) {
					State = SyncState.Deleted
				};
			}
			return remoteItem;
		}

		/// <summary>
		/// <see cref="RemoteProvider.CreateNewSyncItem"/>
		/// </summary>
		public override IRemoteItem CreateNewSyncItem(SyncItemSchema schema) {
			var value = new Exchange.EmailMessage(Service);
			var remoteItem = new ExchangeEmailMessage(schema, value, TimeZone);
			return remoteItem;
		}

		/// <summary>
		/// <see cref="RemoteProvider.KnownTypes"/>
		/// </summary>
		public override IEnumerable<Type> KnownTypes() {
			return _syncListTypes;
		}

		/// <summary>
		/// <see cref="RemoteProvider.GetLocallyModifiedItemsMetadata"/>
		/// </summary>
		/// <remarks>
		/// Email integration only loads messages from Exchange.
		/// </remarks>
		public override IEnumerable<ItemMetadata> GetLocallyModifiedItemsMetadata(SyncContext context,
				EntitySchemaQuery modifiedItemsEsq) {
			return new List<ItemMetadata>();
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

		#endregion

	}

	/// <summary>
	/// Provides user settings for email synchronization provider.
	/// </summary>
	public class EmailExchangeSettings : ExchangeSettings
	{

		#region LocalizableStrings

		/// <summary>
		/// Localized "Mailbox does not exist" error text.
		/// </summary>
		public LocalizableString NoSettingsFoundMessage => new LocalizableString(
			_userConnection.Workspace.ResourceStorage, "EmailExchangeSettings",
			"LocalizableStrings.NoSettingsFoundMessage.Value").ToString();

		#endregion

		#region Constructors: Public

		/// <summary>Creates new <see cref="EmailExchangeSettings"/> instance.</summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		public EmailExchangeSettings(UserConnection userConnection, string senderEmailAddress)
			: base(userConnection, senderEmailAddress) {
		}

		#endregion

		#region Properties: Public

		public ILog Log { get; } = LogManager.GetLogger("Exchange");

		/// <summary>
		/// Current <see cref="MailboxSyncSettings"/> unique identifier.
		/// </summary>
		public Guid MailboxId {
			get;
			set;
		}

		/// <summary>
		/// LastSyncDate minutes offset for filtration emails in synchronization process.
		/// </summary>
		public int LastSyncDateMinutesOffset {
			get;
			set;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Sets <see cref="LastSyncDateMinutesOffset"/> property value.
		/// <see cref="LastSyncDateMinutesOffset"/> must allways be negative integer value, so
		/// all <paramref name="rawValue"/> positive values will be converted to negative values.
		/// </summary>
		/// <param name="rawValue">Current mailbox last synchronization date offset in minutes.</param>
		protected virtual void SetLastSyncDateMinutesOffset(int rawValue) {
			LastSyncDateMinutesOffset = rawValue <= 0 ? rawValue : -1 * rawValue;
		}

		/// <summary>
		/// Sets <see cref="LastSyncDate"/> property value.
		/// Max of <paramref name="loadEmailsFromDate"/> and <paramref name="lastSyncDate"/> 
		/// values used for <see cref="LastSyncDate"/> property value.
		/// When max of parameters values equals <see cref="DateTimeUtilities.JavascriptMinDateTime"/>
		/// then <see cref=" DateTime.MinValue"/> will be used as <see cref="LastSyncDate"/> property value.
		/// </summary>
		/// <param name="loadEmailsFromDate"><see cref="MailboxSyncSettings.LoadEmailsFromDate"/> column value.</param>
		/// <param name="lastSyncDate"><see cref="MailboxSyncSettings.LastSyncDate"/> column value.</param>
		protected virtual void SetLastSyncDate(DateTime loadEmailsFromDate, DateTime lastSyncDate) {
			LastSyncDate = loadEmailsFromDate > lastSyncDate ? loadEmailsFromDate : lastSyncDate;
			Log.Info($"loadEmailsFromDate = '{loadEmailsFromDate}', LastSyncDate = '{LastSyncDate}', result = '{LastSyncDate}'");
			if (LastSyncDate == DateTimeUtilities.JavascriptMinDateTime) {
				LastSyncDate = DateTime.MinValue;
			}
		}

		/// <summary>
		/// Sets remote folders synchronization options.
		/// <see cref="LoadAll"/> property will be set with <paramref name="loadAllEmails"/> value.
		/// <see cref="RemoteFolderUIds"/> property will be set with dictionary which contains remote folrer uid to
		/// <see cref="ActivityFolder"/> unique identifier.
		/// <see cref="RootFolderId"/> property will be set with <paramref name="activityRootFolderId"/> value.
		/// </summary>
		/// <param name="loadAllEmails"><see cref="MailboxSyncSettings.LoadAllEmailsFromMailBox"/> column value.</param>
		/// <param name="activityRootFolderId">Mailbox root <see cref="ActivityFolder"/> unique identifier.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		protected virtual void SetRemoteFoldersOptions(bool loadAllEmails, Guid activityRootFolderId, UserConnection userConnection) {
			LoadAll = loadAllEmails;
			Dictionary<string, Guid> folderUIds = null;
			if (!loadAllEmails) {
				var foldersCorrespondenceEsq = new EntitySchemaQuery(userConnection.EntitySchemaManager,
					"MailboxFoldersCorrespondence");
				foldersCorrespondenceEsq.PrimaryQueryColumn.IsAlwaysSelect = true;
				string uniqueIdColumnName = foldersCorrespondenceEsq.AddColumn("FolderPath").Name;
				string activityFolderIdColumnName = foldersCorrespondenceEsq.AddColumn("ActivityFolder.Id").Name;
				foldersCorrespondenceEsq.Filters.Add(foldersCorrespondenceEsq
					.CreateFilterWithParameters(FilterComparisonType.Equal, "Mailbox", MailboxId));
				var queryResults = foldersCorrespondenceEsq.GetEntityCollection(userConnection);
				folderUIds = queryResults.ToDictionary(result => result.GetTypedColumnValue<string>(uniqueIdColumnName),
					result => result.GetTypedColumnValue<Guid>(activityFolderIdColumnName));
			}
			RemoteFolderUIds = folderUIds ?? new Dictionary<string, Guid>();
			RootFolderId = activityRootFolderId;
		}

		/// <summary>
		/// Sets settings properties.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		protected override void InitializeProperties(UserConnection userConnection) {
			var mssEntitySchemaQuery = new EntitySchemaQuery(userConnection.EntitySchemaManager, "MailboxSyncSettings");
			string primaryColumnName = mssEntitySchemaQuery
				.AddColumn(mssEntitySchemaQuery.RootSchema.GetPrimaryColumnName()).Name;
			mssEntitySchemaQuery.AddColumn("MailSyncPeriod");
			string lastSyncDateColumnName = mssEntitySchemaQuery.AddColumn("LastSyncDate").Name;
			string loadAllEmailsColumnName = mssEntitySchemaQuery.AddColumn("LoadAllEmailsFromMailBox").Name;
			string activityRootFolderIdColumnName = mssEntitySchemaQuery
				.AddColumn("[ActivityFolder:Name:MailboxName].Id").Name;
			string syncDateMinutesOffsetColumnName = mssEntitySchemaQuery.AddColumn("SyncDateMinutesOffset").Name;
			mssEntitySchemaQuery.Filters.Add(mssEntitySchemaQuery.CreateFilterWithParameters(FilterComparisonType.Equal,
				"SenderEmailAddress", SenderEmailAddress));
			mssEntitySchemaQuery.Filters.Add(mssEntitySchemaQuery.CreateFilterWithParameters(FilterComparisonType.Equal,
				"SysAdminUnit", userConnection.CurrentUser.Id));
			var queryResults = mssEntitySchemaQuery.GetEntityCollection(userConnection);
			if (!queryResults.Any()) {
				throw new ApplicationException(string.Format(NoSettingsFoundMessage, SenderEmailAddress));
			}
			Entity syncSetting = queryResults.First();
			var lastSyncDateColumnValue = syncSetting.GetTypedColumnValue<DateTime>(lastSyncDateColumnName);
			var loadAllEmails = syncSetting.GetTypedColumnValue<bool>(loadAllEmailsColumnName);
			var activityRootFolderId = syncSetting.GetTypedColumnValue<Guid>(activityRootFolderIdColumnName);
			DateTime loadFromDate = LoadFromDateType.GetInstance(userConnection)
				.GetConvertLoadFromDate(syncSetting, userConnection.CurrentUser.TimeZone);
			MailboxId = syncSetting.GetTypedColumnValue<Guid>(primaryColumnName);
			SetLastSyncDate(loadFromDate, lastSyncDateColumnValue);
			SetLastSyncDateMinutesOffset(syncSetting.GetTypedColumnValue<int>(syncDateMinutesOffsetColumnName));
			SetRemoteFoldersOptions(loadAllEmails, activityRootFolderId, userConnection);
		}

		#endregion

	}
}
