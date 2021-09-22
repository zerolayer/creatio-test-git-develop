namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Linq;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Terrasoft.Sync;
	using Terrasoft.Configuration;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using SysSettings = Terrasoft.Core.Configuration.SysSettings;
	using AdjustmentRule = System.TimeZoneInfo.AdjustmentRule;

	/// <summary>
	/// Provider for synchronize with Exchange Server.
	/// </summary>
	public abstract class ExchangeSyncProvider : BaseExchangeSyncProvider
	{
		#region Constants: Private

		/// <summary>
		/// Session information template.
		/// </summary>
		private const String _sessionInfoTemplate = "[SyncSessionId: {0}, utc now: {1}, mailbox: {2}]  ";

		/// <summary>
		/// Session identifiers.
		/// </summary>
		private readonly long _sessionId = DateTime.Now.Ticks;

		#endregion

		#region Constants: Protected

		protected const int PageItemCount = 41;

		#endregion

		#region Fields: Private

		private readonly Guid _storeId;

		#endregion

		#region Constructors: Protected

		protected ExchangeSyncProvider(Guid storeId, TimeZoneInfo timeZone, string senderEmailAddress)
			: base(new ExchangeUtilityImpl().Log) {
			_storeId = storeId;
			TimeZone = timeZone;
			SenderEmailAddress = senderEmailAddress;
		}

		#endregion

		#region Properties: Protected

		protected Exchange.ExchangeService Service {
			get;
			set;
		}

		/// <summary>
		/// Current time zone of BPMonline user.
		/// </summary>
		/// <remarks>
		/// It converts a BPMonline local time server
		/// to the user's local time when using the local storage and 
		/// converts local user time to a local time server BPMonline, 
		/// when using remote repository.
		/// </remarks>
		protected TimeZoneInfo TimeZone {
			get;
			private set;
		}

		protected ExchangeUtilityImpl ExchangeUtility { get; } = new ExchangeUtilityImpl();

		#endregion

		#region Properties: Public

		/// <summary>
		/// Unique identifier in remote storage.
		/// </summary>
		public override sealed Guid StoreId {
			get {
				return _storeId;
			}
		}

		/// <summary>
		///  Sender email address
		/// </summary>
		public string SenderEmailAddress {
			get;
			private set;
		}

		/// <summary>
		/// Session information.
		/// </summary>
		public override string SessionInfo {
			get {
				return String.Format(_sessionInfoTemplate, _sessionId, DateTime.UtcNow, SenderEmailAddress);
			}
		}

		#endregion

		#region Methods: Private

		/// <summary> Updates extra parameters in exchange appointment item.</summary>
		/// <param name="syncValueName">The name of the entity type.</param>
		private void UpdateAppointmentExtraParameters(string syncValueName, IRemoteItem syncItem) {
			if (syncValueName == ExchangeConsts.ExchangeAppointmentClassName) {
				var exchangeAppointmentSyncItem = (ExchangeAppointment)syncItem;
				var exchangeAppointment = (Exchange.Appointment)exchangeAppointmentSyncItem.Item;
				var propertySet = new Exchange.PropertySet(Exchange.AppointmentSchema.ICalUid);
				propertySet.AddRange(new[] {
					Exchange.ItemSchema.Subject,
					Exchange.ItemSchema.Sensitivity
				});
				exchangeAppointment.Load(propertySet);
				exchangeAppointmentSyncItem.UpdateExtraParameters(exchangeAppointment);
			};
		}

		/// <summary>
		/// Returns Id of first child folder by name and root folder id.
		/// If the value can not be obtained, the method returns <c>null</c>.</summary>
		/// <param name="exchangeService">Exchange service.</param>
		/// <param name="parentId">Root folder Id.</param>
		/// <param name="childName">Display name of the finding folder.</param>
		/// <returns>Id of child folder, or <c>null</c>.</returns>
		private Exchange.FolderId GetChildIdByName(Exchange.ExchangeService exchangeService,
			Exchange.FolderId parentId, string childName) {
			Exchange.SearchFilter filter = new Exchange.SearchFilter
				.IsEqualTo(Exchange.FolderSchema.DisplayName, childName);
			var result = exchangeService.FindFolders(parentId, filter, new Exchange.FolderView(1));
			if (!result.Folders.Any()) {
				return null;
			}
			return result.Folders[0].Id;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Fills list exchange folders.
		/// </summary>
		/// <param name="folders">List <see cref="Exchange.Folder"/>.</param>
		/// <param name="context"><see cref="SyncContext"/>instance.</param>
		protected virtual void FillRemoteFoldersList(ref List<Exchange.Folder> folders, SyncContext context) {
		}

		/// <summary>
		/// Fills Exchange item filter collection.
		/// </summary>
		protected virtual void FillItemsFilterCollection() {
		}

		/// <summary>
		/// Gets list remote items.
		/// </summary>
		/// <returns>List remote items.</returns>
		protected virtual List<IRemoteItem> GetItemsFromFolders() {
			return new List<IRemoteItem>();
		}

		/// <summary>
		/// Initializes Exchange service.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <returns>New Exchange service instance.</returns>
		protected virtual Exchange.ExchangeService InitializeService(UserConnection userConnection) {
			return ExchangeUtility.CreateExchangeService(userConnection, SenderEmailAddress);
		}

		/// <summary>
		/// Checks if Appointment includes other participants in addition to the organizer.
		/// </summary>
		/// <param name="exchangeItem">Exchange appointment instance.</param>
		/// <param name="syncItem">Remote storage item.</param>
		protected virtual bool HasAnotherAttendees(Exchange.Appointment exchangeItem) {
			string organizerAddress = exchangeItem.Organizer.Address;
			bool hasAnotherAttendees = false;
			foreach (Exchange.Attendee attendee in exchangeItem.RequiredAttendees) {
				if (!string.Equals(organizerAddress, attendee.Address, StringComparison.OrdinalIgnoreCase)) {
					hasAnotherAttendees = true;
					break;
				}
			}
			return hasAnotherAttendees;
		}

		/// <summary>
		/// Removes deprecated for loading folders from collection
		/// </summary>
		/// <param name="folders">List of folders to fe filtered</param>
		protected virtual void FilterDeprecatedFolders(ref List<Exchange.Folder> folders) {
			var id = new Exchange.FolderId(Exchange.WellKnownFolderName.DeletedItems, SenderEmailAddress);
			Exchange.FolderId[] deprecatedIds = {
				GetConflictsFolderId(Service),
				Exchange.Folder.Bind(Service, id).Id
			};
			folders.RemoveAll(x => deprecatedIds.Contains(x.Id));
		}

		/// <summary>
		/// Gets unique folder identifier.
		/// </summary>
		/// <param name="folder"><see cref="Exchange.Folder"/></param>
		/// <returns>Unique folder identifier</returns>
		protected virtual string GetFolderId(Exchange.Folder folder) {
			return folder.Id.UniqueId;
		}

		/// <summary>
		/// Attempts to bind <see cref="Exchange.Folder"/> items by <paramref name="folderRemoteIds"/> collection.
		/// The method returns only folders which can not be obtained.
		/// </summary>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="folderRemoteIds">Exchange folders remote ids collection.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns>Folders which can not be obtained.</returns>
		protected virtual List<Exchange.Folder> SafeBindFolders(Exchange.ExchangeService service,
				IEnumerable<string> folderRemoteIds, SyncContext context) {
			List<Exchange.Folder> result = new List<Exchange.Folder>();
			foreach (string uniqueId in folderRemoteIds) {
				if (string.IsNullOrEmpty(uniqueId)) {
					continue;
				}
				try {
					Exchange.Folder folder = Exchange.Folder.Bind(service, new Exchange.FolderId(uniqueId));
					result.Add(folder);
				} catch (Exchange.ServiceResponseException exception) {
					LogMessage(context, SyncAction.None, SyncDirection.Upload, "Error while folder bind. Folder remoteId {0}", exception, uniqueId);
				}
			}
			return result;
		}

		/// <summary>
		/// Checks is current sync action ignored by this synchronization.
		/// <see cref="SyncAction.Delete"/> ignored by default.
		/// </summary>
		/// <param name="action">Current <see cref="SyncAction"/>value.</param>
		/// <returns><c>True</c> if <paramref name="action"/> ignored, <c>false</c> otherwise.</returns>
		protected virtual bool GetCurrentActionIgnored(SyncAction action) {
			return action == SyncAction.Delete;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Gets Items from <see cref="folder"/> by <see cref="filterCollection"/>
		/// </summary>
		/// <param name="folder">Exchange folder.</param>
		/// <param name="filterCollection">Filter collection.</param>
		/// <param name="itemView">Represents the view settings in a folder search operation.</param>
		/// <returns></returns>
		public virtual Exchange.FindItemsResults<Exchange.Item> GetFolderItemsByFilter(Exchange.Folder folder,
			Exchange.SearchFilter.SearchFilterCollection filterCollection, Exchange.ItemView itemView) {
			return folder.ReadItems(filterCollection, itemView);
		}

		/// <summary>
		/// Attempts to obtain a copy of this type by value identifier <paramref name="itemId"/>.
		/// If the value can not be obtained, the method returns <c>null</c>.
		/// </summary>
		/// <typeparam name="T">Type, inheritor <see cref="Exchange.Item"/>.</typeparam>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="itemId">Represents the Id of an Exchange item.</param>
		/// <returns>An instance of this type, or <c>null</c>.</returns>
		public virtual T SafeBindItem<T>(Exchange.ExchangeService service, Exchange.ItemId itemId)
			where T : Exchange.Item {
			return ExchangeUtility.SafeBindItem<T>(Service, itemId);
		}

		/// <summary>
		/// Returns the external data repository synchronization object by a unique foreign key.
		/// </summary>
		/// <param name="schema">Schema.</param>
		/// <param name="id">Unique foreign key.</param>
		/// <returns>Created instance of <see cref="IRemoteItem"/>.</returns>
		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, string id) {
			return null;
		}

		/// <summary>
		/// Executes CRUD operations for remote storage item.
		/// </summary>
		/// <param name="context">Sync context.</param>
		/// <param name="syncItem">Remote storage item.</param>
		public override void ApplyChanges(SyncContext context, IRemoteItem syncItem) {
			SyncAction action = syncItem.Action;
			if (action == SyncAction.None || GetCurrentActionIgnored(action)) {
				return;
			}
			var exchangeSyncItem = (ExchangeBase)syncItem;
			Exchange.Item exchangeItem = exchangeSyncItem.Item;
			string displayName = "";
			string syncValueName = syncItem.Schema.SyncValueName;
			bool dontSendNotifications = syncValueName == ExchangeConsts.ExchangeAppointmentClassName;
			try {
				switch (action) {
					case SyncAction.Create:
						displayName = exchangeSyncItem.DisplayName;
						if (dontSendNotifications) {
							var exchangeAppointment = ((Exchange.Appointment)exchangeItem);
							exchangeAppointment.Save(Exchange.SendInvitationsMode.SendToNone);
						} else {
							exchangeItem.Save();
						}
						UpdateAppointmentExtraParameters(syncValueName, syncItem);
						break;
					case SyncAction.Update:
						displayName = exchangeSyncItem.DisplayName;
						if (dontSendNotifications) {
							((Exchange.Appointment)exchangeItem).Update(Exchange.ConflictResolutionMode.AlwaysOverwrite,
								Exchange.SendInvitationsOrCancellationsMode.SendToNone);
						} else {
							exchangeItem.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite);
						}
						break;
					default:
						if (exchangeSyncItem.State != SyncState.Deleted) {
							if (dontSendNotifications) {
								((Exchange.Appointment)exchangeItem).Delete(Exchange.DeleteMode.MoveToDeletedItems,
									Exchange.SendCancellationsMode.SendToNone);
							} else {
								exchangeItem.Delete(Exchange.DeleteMode.MoveToDeletedItems);
							}
							exchangeSyncItem.State = SyncState.Deleted;
						}
						break;
				}
				LogMessage(context, action, SyncDirection.Upload, OperationInfoLczStrings[action], null, displayName,
					syncValueName);
			} catch (Exception ex) {
				LogMessage(context, action, SyncDirection.Upload, OperationInfoLczStrings[action], ex, displayName,
					syncValueName);
			}
		}

		/// <summary>
		/// Writes log message.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="action"><see cref="SyncAction"/> enumeration.</param>
		/// <param name="syncDirection"><see cref="SyncDirection"/> enumeration.</param>
		/// <param name="format">String a composite format.</param>
		/// <param name="ex">Exception.</param>
		/// <param name="args">An array of additional parameters.</param>
		public virtual void LogMessage(SyncContext context, SyncAction action, SyncDirection syncDirection, string format, Exception ex, params object[] args) {
			if (ex == null) {
				context.LogInfo(action, syncDirection, format, args);
			} else {
				context.LogError(action, syncDirection, format, ex, args);
			}
		}

		/// <summary>
		/// Returns enumerator for synchronization objects of remote storage.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <returns>The enumerator object synchronization of external storage.
		/// </returns>
		/// <remarks> Calls a Exchange service initialization.</remarks>
		public override IEnumerable<IRemoteItem> EnumerateChanges(SyncContext context) {
			Service = Service ?? InitializeService(context.UserConnection);
			return null;
		}

		/// <summary>
		/// Returns true if <see cref="Microsoft.Exchange.WebServices.Data.Appointment"/> must be saved
		/// without participants notifications.</summary>
		/// <param name="exchangeAppointment">Exchange appointment instance.</param>
		/// <param name="exchangeItem">Exchange item instance.</param>
		/// <returns>Dont send notifications flag.</returns>
		public bool GetDontSendNotifications(ExchangeAppointment exchangeAppointment, Exchange.Appointment exchangeItem) {
			SyncAction action = exchangeAppointment.Action;
			return !(exchangeAppointment.SendNotifications &&
				(action == SyncAction.Create || HasAnotherAttendees(exchangeItem))
			);
		}

		/// <summary>
		/// Attempts to get Id of Conflicts folder
		/// If the value can not be obtained, the method returns <c>null</c>.</summary>
		/// <param name="exchangeService">Exchange service.</param>
		/// <returns>Id of this folder, or <c>null</c>.</returns>
		public Exchange.FolderId GetConflictsFolderId(Exchange.ExchangeService exchangeService) {
			var id = new Exchange.FolderId(Exchange.WellKnownFolderName.MsgFolderRoot, SenderEmailAddress);
			var rootFolderId = Exchange.Folder.Bind(exchangeService, id).Id;
			Exchange.FolderId syncIssuesFolderId = GetChildIdByName(exchangeService, rootFolderId, "Sync Issues");
			if (syncIssuesFolderId == null) {
				return null;
			}
			return GetChildIdByName(exchangeService, syncIssuesFolderId, "Conflicts");
		}

		#endregion

		}

}