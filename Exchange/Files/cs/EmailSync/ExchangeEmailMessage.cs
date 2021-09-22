namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using Microsoft.Exchange.WebServices.Data;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Common;
	using Terrasoft.Common.Json;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Entities;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.Mail.Sender;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeEmailMessage

	/// <summary>
	/// Exchange email message remote item class.
	/// </summary>
	[Map("Activity", 0, IsPrimarySchema = true, Direction = SyncDirection.Download)]
	[Map("ActivityFile", 1, PrimarySchemaName = "Activity", ForeingKeyColumnName = "Activity",
		Direction = SyncDirection.Download)]
	[Map("ActivityInFolder", 2, PrimarySchemaName = "Activity", ForeingKeyColumnName = "Activity",
		Direction = SyncDirection.Download)]
	public class ExchangeEmailMessage : ExchangeBase
	{

		#region Consts: Private

		private const string _emailAttachmentExtension = ".eml";

		private const string _contactAttachmentExtension = ".vcf";

		private const string _appoitmentAttachmentExtension = ".ics";

		#endregion

		#region Fields: Private

		private Exchange.PropertySet _propertySet;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Creates new <see cref="ExchangeEmailMessage"/> instance,
		/// using <paramref name"schema"/> and <paramref name="item"/>.
		/// </summary>
		/// <param name="schema">Sync element columns and objects map.</param>
		/// <param name="item">Exchange storage element instance.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeEmailMessage(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timeZoneInfo)
			: base(schema, item, timeZoneInfo) {
			ActivityFolderIds = new List<Guid>();
			SetPropertySet();
		}

		/// <summary>
		/// Creates new <see cref="ExchangeEmailMessage"/> instance,
		/// using <paramref name"schema"/> and <paramref name="item"/> and <paramref name="remoteId"/> used as RemoteId.
		/// </summary>
		/// <param name="schema">Sync element columns and objects map.</param>
		/// <param name="item">Exchange storage element instance.</param>
		/// <param name="remoteId">Unique remote item id.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeEmailMessage(SyncItemSchema schema, Exchange.Item item, string remoteId,
			TimeZoneInfo timeZoneInfo)
			: this(schema, item, timeZoneInfo) {
			RemoteId = remoteId;
			SetPropertySet();
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// <see cref="ActivityFolder"/> ids.
		/// </summary>
		public List<Guid> ActivityFolderIds {
			get;
			set;
		}

		/// <summary>
		/// Current exchange email message instance.
		/// </summary>
		private Exchange.EmailMessage _emailMessage;
		public Exchange.EmailMessage EmailMessage {
			get {
				return _emailMessage ?? (_emailMessage = Item as Exchange.EmailMessage);
			}
		}

		#endregion

		#region Methods: Private

		private void SetPropertySet() {
			_propertySet = new Exchange.PropertySet(Exchange.BasePropertySet.FirstClassProperties,
				ExchangeUtility.GetContactExtendedPropertyDefinition());
		}

		/// <summary>
		/// Returns true if local item contains Activity instance.
		/// </summary>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="activityId">Activity instance id.</param>
		/// <returns>
		/// True if local item contains Activity instance.
		/// </returns>
		private bool IsMessageSyncedInBpm(LocalItem localItem, Guid activityId) {
			var collection = localItem.Entities["Activity"];
			if (collection.Count == 0) {
				return false;
			}
			var activitySyncEntity = collection.FirstOrDefault(e => e.EntityId == activityId);
			return (activitySyncEntity != null && activitySyncEntity.Action != SyncAction.Create);
		}

		/// <summary>
		/// Validates email message.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		private bool IsMessageValid(SyncContext context) {
			if (EmailMessage == null) {
				context.LogError(Action, SyncDirection.Upload, "Item {0} has invalid type", Item.Subject);
				return false;
			} else if (!IsIncommingMessage(context, EmailMessage)) {
				return false;
			}
			return true;
		}

		/// <summary>
		/// Creates <see cref="SyncEntity"/> instance.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="message">Exchange email message.</param>
		/// <param name="subject">Email message subject.</param>
		/// <param name="activity">Activity instance.</param>
		/// <returns><see cref="SyncEntity"/> instance.</returns>
		private SyncEntity GetActivityInstanceSync(SyncContext context, Exchange.EmailMessage message, string subject,
				Activity activity) {
			var emailService = ClassFactory.Get<IEmailService>(new ConstructorArgument("uc", context.UserConnection));
			var emailIds = emailService.GetActivityIds(EmailMessage.InternetMessageId);
			if (emailIds == null || !emailIds.Any()) {
				DateTime sendDate = GetMessageDateProperty(message, Exchange.ItemSchema.DateTimeSent, context.UserConnection);
				emailIds = ActivityUtils.GetExistingEmaisIds(context.UserConnection, sendDate, subject,
					message.Body.Text, context.UserConnection.CurrentUser.TimeZone);
			}
			bool activityExists = emailIds.Any();
			if (activityExists) {
				activity.PrimaryColumnValue = emailIds.FirstOrDefault();
			} else {
				activity.SetDefColumnValues();
			}
			SyncAction action = activityExists ? SyncAction.None : SyncAction.Create;
			return new SyncEntity(activity, SyncState.None) {
				Action = action
			};
		}

		/// <summary>
		/// Fills email attachment instance.
		/// </summary>
		/// <param name="activityFile">Activcity file instance.</param>
		/// <param name="attachment">Exchange email message attachment.</param>
		/// <param name="mainRecordId">Activity instance id.</param>
		private void FillActivityFile(ref Entity activityFile, Attachment attachment, UserConnection userConnection, Guid mainRecordId) {
			string contentId = attachment.ContentId;
			bool inline = attachment.IsInline ||
				(contentId.IsNotNullOrEmpty() &&
				EmailMessage.Body.Text.IsNotNullOrEmpty() &&
				EmailMessage.Body.Text.Contains("cid:" + contentId));
			activityFile.SetColumnValue("Name", GetAttachmentName(userConnection, attachment));
			activityFile.SetColumnValue("Version", 1);
			activityFile.SetColumnValue("TypeId", Terrasoft.WebApp.FileConsts.FileTypeUId);
			activityFile.SetColumnValue("ActivityId", mainRecordId);
			activityFile.SetColumnValue("ModifiedOn", attachment.LastModifiedTime);
			activityFile.SetColumnValue("Uploaded", false);
			activityFile.SetColumnValue("Inline", inline);
			var externalStorageProperties = new JObject {
				["RemoteId"] = attachment.Id,
				["ActivityId"] = mainRecordId.ToString(),
				["ExchangeMessageId"] = Id
			};
			activityFile.SetColumnValue("ExternalStorageProperties", Json.Serialize(externalStorageProperties));
		}

		/// <summary>
		/// Loads attachment item and returns attachment file name with or without extension, depending on the attachment type.
		/// </summary>
		/// <param name="attachment">Exchange email message attachment.</param>
		/// <returns>Email attachment name.</returns>
		private string GetAttachmentName(UserConnection userConnection, Attachment attachment) {
			string attachmentName = attachment.Name;
			if (attachment is ItemAttachment) {
				var itemAttachment = attachment as ItemAttachment;
				itemAttachment.Load(new PropertySet(BasePropertySet.IdOnly));
				var item = itemAttachment.Item;
				switch (item.ItemClass) {
					case "IPM.Note":
						attachmentName += _emailAttachmentExtension;
						break;
					case "IPM.Appointment":
						attachmentName += _appoitmentAttachmentExtension;
						break;
					case "IPM.Contact":
						attachmentName += _contactAttachmentExtension;
						break;
				}
			}
			return ActivityUtils.GetAttachmentName(userConnection, attachmentName);
		}

		/// <summary>
		/// Creates ActivityFile instance for email message attachment.
		/// </summary>
		/// <param name="userConnection">User connection instance.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="attachment">Exchange email message attachment.</param>
		/// <param name="mainRecordId">Activity instance id.</param>
		/// <returns>ActivityFile instance.</returns>
		private Entity AddNewAttachment(UserConnection userConnection, LocalItem localItem,
				Attachment attachment, Guid mainRecordId) {
			EntitySchema schema = userConnection.EntitySchemaManager.GetInstanceByName("ActivityFile");
			var activityFile = schema.CreateEntity(userConnection);
			activityFile.SetDefColumnValues();
			SyncEntity syncEntity = SyncEntity.CreateNew(activityFile);
			if (!userConnection.GetIsFeatureEnabled("DoNotUseMetadataForEmail")) {
				var attachmentExtendProperty = new JObject();
				attachmentExtendProperty["RemoteId"] = attachment.Id;
				attachmentExtendProperty["ActivityId"] = mainRecordId.ToString();
				ExchangeUtility.SetActivityExtraParameters(syncEntity, attachmentExtendProperty);
			}
			localItem.AddOrReplace("ActivityFile", syncEntity);
			FillActivityFile(ref activityFile, attachment, userConnection, mainRecordId);
			return activityFile;
		}

		/// <summary>
		/// Fills email attachment detail.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="mainRecordId">Activity instance id.</param>
		private void FillAttachmentsDetail(SyncContext context, LocalItem localItem, Guid mainRecordId) {
			LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] setting activity {0} attachments from message {1}.",
				mainRecordId, GetDisplayName());
			UserConnection userConnection = context.UserConnection;
			if (EmailMessage == null || IsMessageSyncedInBpm(localItem, mainRecordId)) {
				string logMessageTpl = EmailMessage == null
					? "[FillAttachmentsDetail] Email message is null"
					: "[FillAttachmentsDetail] Message already synced to bpm (Activity Id: {0})";
				LogInfo(context, Action, SyncDirection.Upload, logMessageTpl, mainRecordId.ToString());
				return;
			}
			IEnumerable<SyncEntity> filesToDelete = localItem.Entities["ActivityFile"].Where(e =>
				e.State != SyncState.New && IsMessageNotContainsAttachment(e, EmailMessage, mainRecordId));
			foreach (SyncEntity syncEntity in filesToDelete) {
				syncEntity.Action = SyncAction.Delete;
			}
			LogInfo(context, Action, SyncDirection.Upload, "[FillAttachmentsDetail] email attachments count: {0}", EmailMessage.Attachments.Count);
			ProcessAttachments(context, localItem, mainRecordId, EmailMessage.Attachments);
		}

		/// <summary>
		/// Creates activity files based on email attachmets and saves it to database.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="mainRecordId">Activity instance id.</param>
		/// <param name="emailMessageAttachments">Email attachment collection.</param>
		private void ProcessAttachments(SyncContext context, LocalItem localItem, Guid mainRecordId, AttachmentCollection emailMessageAttachments) {
			foreach (Attachment attachment in emailMessageAttachments) {
				Attachment emailAttachment = attachment as FileAttachment;
				if (emailAttachment == null) {
					emailAttachment = attachment as ItemAttachment;
				}
				if (emailAttachment == null) {
					continue;
				}
				UserConnection userConnection = context.UserConnection;
				Entity activityFile = null;
				IEnumerable<SyncEntity> oldFiles =
					localItem.Entities["ActivityFile"].Where(
						e => e.State != SyncState.New && IsAttachmentSyncEntityValid(e, emailAttachment.Id, mainRecordId));
				var syncEntities = oldFiles as IList<SyncEntity> ?? oldFiles.ToList();
				if (syncEntities.Any()) {
					activityFile = syncEntities.First().Entity;
					if (activityFile.GetTypedColumnValue<DateTime>("ModifiedOn") < emailAttachment.LastModifiedTime) {
						FillActivityFile(ref activityFile, emailAttachment, userConnection, mainRecordId);
					}
				} else {
					activityFile = AddNewAttachment(userConnection, localItem, emailAttachment, mainRecordId);
					LogInfo(context, Action, SyncDirection.Upload, "[FillAttachmentsDetail] created attachment {0} Id: {1}",
						activityFile.GetTypedColumnValue<string>("Name"), activityFile.PrimaryColumnValue.ToString());
				}
				if (activityFile.GetTypedColumnValue<bool>("Inline")) {
					var url = string.Format("../rest/FileService/GetFile/{0}/{1}", activityFile.Schema.UId, activityFile.PrimaryColumnValue);
					LogInfo(context, Action, SyncDirection.Upload,
						"[FillAttachmentsDetail] attachment {0} is inline, replacing cid {1} with url {2}", activityFile.GetTypedColumnValue<string>("Name"),
						emailAttachment.ContentId, url);
					EmailMessage.Body.Text = EmailMessage.Body.Text.Replace("cid:" + emailAttachment.ContentId, url);
				}
			}
		}

		/// <summary>
		/// Generates ActivityInFolder list for delete.
		/// Entries for removal consist of those that do not exist in the table ActivityInFolder, 
		/// but written in SysSyncMetaData and those that have been modified, but not listed as a sync folder.
		/// </summary>
		/// <param name="localItem">Local storage item.</param>
		private IEnumerable<SyncEntity> GetActivityInFolderForDelete(LocalItem localItem) {
			List<SyncEntity> foldersToDelete = new List<SyncEntity>();
			IEnumerable<SyncEntity> notLoadedFolders = localItem.Entities["ActivityInFolder"].Where(e =>
				e.Entity.LoadState != EntityLoadState.Loaded && e.State != SyncState.New);
			IEnumerable<SyncEntity> notSyncededFolders = localItem.Entities["ActivityInFolder"].Where(e =>
				e.State != SyncState.New && e.Entity.LoadState == EntityLoadState.Loaded &&
				!ActivityFolderIds.Contains(e.Entity.GetTypedColumnValue<Guid>("FolderId")));
			if (notLoadedFolders != null) {
				foldersToDelete.AddRange(notLoadedFolders);
			}
			if (notSyncededFolders != null) {
				foldersToDelete.AddRange(notSyncededFolders);
			}
			return foldersToDelete;
		}

		/// <summary>
		/// Fills ActivityInFolder table for email.
		/// </summary>
		/// <param name="userConnection">User connection instance.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="syncEntity"><see cref="SyncEntity"/> instance.</param>
		private void FillActivityInFolders(UserConnection userConnection, LocalItem localItem, SyncEntity syncEntity) {
			if (syncEntity != null && syncEntity.Action != SyncAction.Create) {
				return;
			}
			Guid mainRecordId = syncEntity.EntityId;
			IEnumerable<SyncEntity> foldersToDelete = GetActivityInFolderForDelete(localItem);
			foreach (SyncEntity folderSyncEntity in foldersToDelete) {
				folderSyncEntity.Action = SyncAction.Delete;
			}
			foreach (var activityFolder in ActivityFolderIds) {
				Guid folder = activityFolder;
				IEnumerable<SyncEntity> oldFolder = localItem.Entities["ActivityInFolder"].Where(e =>
					e.State != SyncState.New &&
					e.Entity.LoadState == EntityLoadState.Loaded &&
					e.Entity.GetTypedColumnValue<Guid>("FolderId") == folder &&
					e.Entity.GetTypedColumnValue<Guid>("ActivityId") == mainRecordId);
				var syncEntities = oldFolder as SyncEntity[] ?? oldFolder.ToArray();
				if (syncEntities.Any()) {
					continue;
				}
				EntitySchema schema = userConnection.EntitySchemaManager.GetInstanceByName("ActivityInFolder");
				var activityInFolder = (ActivityInFolder)schema.CreateEntity(userConnection);
				activityInFolder.SetDefColumnValues();
				localItem.AddOrReplace("ActivityInFolder", SyncEntity.CreateNew(activityInFolder));
				activityInFolder.FolderId = activityFolder;
				activityInFolder.ActivityId = mainRecordId;
			}
		}

		/// <summary>
		/// Returns email owner contact id.
		/// </summary>
		/// <param name="message">Exchange email message.</param>
		/// <param name="userConnection">User connection instance.</param>
		/// <returns>Email owner contact id.</returns>
		private Guid GetActivityOwnerId(Exchange.EmailMessage message, UserConnection userConnection) {
			Guid contactId = userConnection.CurrentUser.ContactId;
			List<object> contactsEmailList = (
					from emailAddress in message.ToRecipients
					where !string.IsNullOrEmpty(emailAddress.Address)
					select emailAddress.Address)
				.Cast<object>().ToList();
			contactsEmailList.AddRange((
				from emailAddress in message.CcRecipients
				where !string.IsNullOrEmpty(emailAddress.Address)
				select emailAddress.Address));
			if (!contactsEmailList.Any()) {
				return contactId;
			}
			var sysAdminUnitSubSelect = new Select(userConnection)
				.Column("SysAdminUnit", "Id")
				.From("SysAdminUnit")
				.Where("SysAdminUnit", "AccountId").IsEqual("Contact", "AccountId")
				.And("SysAdminUnit", "Id").Not().IsNull();
			var contactSubSelect = new Select(userConnection)
				.Column("Contact", "Id")
				.From("Contact")
				.Where("Contact", "Id").IsEqual("ContactCommunication", "ContactId")
				.And().Exists(sysAdminUnitSubSelect);
			var select = new Select(userConnection).Top(1)
				.Column("ContactCommunication", "ContactId").As("ContactId")
				.From("ContactCommunication")
				.Where("ContactCommunication", "CommunicationTypeId")
				.IsEqual(Column.Parameter(Guid.Parse(CommunicationTypeConsts.EmailId)))
				.And("ContactCommunication", "Number").In(Column.Parameters(contactsEmailList))
				.And().Exists(contactSubSelect) as Select;
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (var reader = select.ExecuteReader(dbExecutor)) {
					if (reader.Read()) {
						var value = reader.GetColumnValue<Guid>("ContactId");
						contactId = value == Guid.Empty ? contactId : value;
					}
				}
			}
			return contactId;
		}

		/// <summary>
		/// Checks if <paramref name="attachmentMetaData"/> is one of <paramref name="message"/> attachments.
		/// </summary>
		/// <param name="attachmentMetaData"><see cref="SyncEntity"/> instance.</param>
		/// <param name="message"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="activityId">Activity instance id.</param>
		/// <returns><c>True</c>, if <paramref name="attachmentMetaData"/> is one of <paramref name="message"/> attachments.</returns>
		private bool IsMessageNotContainsAttachment(SyncEntity attachmentMetaData, Exchange.EmailMessage message, Guid activityId) {
			string attachRemoteId;
			Guid storedActivityId = Guid.Empty;
			ParseAttachmentExtraParams(attachmentMetaData.ExtraParameters, out attachRemoteId, out storedActivityId);
			return (message.Attachments.All(at => at.Id != attachRemoteId) && activityId.Equals(storedActivityId));
		}

		/// <summary>
		/// Returns true if <paramref name="attachmentMetaData"/> ExtraParemeters contains <paramref name="activityId"/>
		/// and <paramref name="attachmentId"/>.
		/// </summary>
		/// <param name="attachmentMetaData"><see cref="SyncEntity"/> instance.</param>
		/// <param name="attachmentId"><see cref="Attachment"/> remote id.</param>
		/// <param name="activityId">Activity instance id.</param>
		/// <returns><c>True</c>, if <paramref name="attachmentMetaData"/> ExtraParemeters contains <paramref name="activityId"/>
		/// and <paramref name="attachmentId"/>.</returns>
		private bool IsAttachmentSyncEntityValid(SyncEntity attachmentMetaData, string attachmentId, Guid activityId) {
			string attachRemoteId;
			Guid storedActivityId = Guid.Empty;
			ParseAttachmentExtraParams(attachmentMetaData.ExtraParameters, out attachRemoteId, out storedActivityId);
			return (attachmentId == attachRemoteId && activityId.Equals(storedActivityId));
		}

		/// <summary>
		/// Extracts attachment remote id and activity id from <paramref name="extraParameters"/> string.
		/// </summary>
		/// <param name="extraParameters">Serialized json string.</param>
		/// <param name="remoteId">Attachment remote id.</param>
		/// <param name="activityId">Activity instance id.</param>
		private void ParseAttachmentExtraParams(string extraParameters, out string remoteId, out Guid activityId) {
			remoteId = ExchangeUtility.TryGetPropertyFromJson(extraParameters, "RemoteId");
			string rawStoredActivityId = ExchangeUtility.TryGetPropertyFromJson(extraParameters, "ActivityId");
			Guid storedActivityId = Guid.Empty;
			Guid.TryParse(rawStoredActivityId, out storedActivityId);
			activityId = storedActivityId;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Calls <see cref="Exchange.Item.Load(Exchange.PropertySet)"/> method for <paramref name="exchangeMessage"/>.
		/// </summary>
		/// <param name="exchangeMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual void LoadItemProperties(Exchange.EmailMessage exchangeMessage) {
			exchangeMessage.Load(_propertySet);
		}

		/// <summary>
		/// Returns DateTime value from <paramref name="propertyDefinition"/> exchange message property.
		/// </summary>
		/// <param name="exchangeMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="propertyDefinition"><see cref="Exchange.PropertyDefinition"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>
		/// DateTime value from <paramref name="propertyDefinition"/>.
		/// </returns>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual DateTime GetMessageDateProperty(Exchange.EmailMessage exchangeMessage,
				Exchange.PropertyDefinition propertyDefinition, UserConnection userConnection) {
			return exchangeMessage.SafeGetValue<DateTime>(propertyDefinition).GetUserDateTime(userConnection);
		}

		/// <summary>
		/// Checks if message is incoming.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="emailMessage">Exchange email message.</param>
		/// <returns>True if message is incoming.</returns>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual bool IsIncommingMessage(SyncContext context, Exchange.EmailMessage emailMessage) {
			LoadItemProperties(emailMessage);
			return true;
		}

		/// <summary>
		/// Calls <see cref="ActivityUtils.SetEmailRelations"/> method tor <paramref name="email"/>.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="email">Activity instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual void UpdateEmailRelations(UserConnection userConnection, Activity email) {
			ActivityUtils.SetEmailRelations(userConnection, email);
		}

		/// <summary>
		/// Creates <see cref="EmailMessageData"/> instance for <paramref name="activity"/> in current mailbox.
		/// </summary>
		/// <param name="activity"><see cref="Entity"/> instance.</param>
		/// <param name="remoteProvider"><see cref="ExchangeEmailSyncProvider"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <remarks>If <paramref name="localItem"/> does not contain entity config for <see cref="EmailMessageData"/>
		/// it will be created.</remarks>
		protected void CreateEmailMessageData(Entity activity, ExchangeEmailSyncProvider remoteProvider,
				LocalItem localItem, UserConnection userConnection) {
			if (localItem.Schema.Configs.All(c => c.SchemaName != "EmailMessageData")) {
				EntityConfig emailMessageDataConfig = new DetailEntityConfig() {
					SchemaName = "EmailMessageData",
					PrimarySchemaName = "Activity",
					ForeingKeyColumnName = "Activity",
					Order = 3
				};
				localItem.Schema.Configs.Add(emailMessageDataConfig);
				localItem.Entities["EmailMessageData"] = new List<SyncEntity>();
			}
			var ticks = ActivityUtils.GetSendDateTicks(userConnection, activity);
			var userConnectionParam = new ConstructorArgument("userConnection", userConnection);
			EmailMessageHelper helper = ClassFactory.Get<EmailMessageHelper>(userConnectionParam);
			Dictionary<string, string> headers = new Dictionary<string, string>() {
				{ "MessageId", EmailMessage.InternetMessageId },
				{ "InReplyTo", EmailMessage.InReplyTo },
				{ "SyncSessionId", remoteProvider.SynsSessionId },
				{ "References", EmailMessage.References },
				{ "SendDateTicks", ticks.ToString() }
			};
			Guid mailboxSyncSettingsId = remoteProvider.UserSettings.MailboxId;
			Entity emailMessageData = helper.CreateEmailMessage(activity, mailboxSyncSettingsId, headers, false);
			if (emailMessageData != null) {
				SyncEntity syncEntity = SyncEntity.CreateNew(emailMessageData);
				localItem.AddOrReplace("EmailMessageData", syncEntity);
			}
		}

		/// <summary> 
		/// Writes information message to the log.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="operation"> Log action for an object synchronization.</param>
		/// <param name="direction">Synchronization direction.</param>
		/// <param name="format">Format.</param>
		/// <param name="args">Format patameters.</param>
		protected override void LogInfo(SyncContext context, SyncAction operation, SyncDirection direction,
			string format, params object[] args) {
			ExchangeEmailSyncProvider provider = context.RemoteProvider as ExchangeEmailSyncProvider;
			string senderEmailAddress = provider.UserSettings.SenderEmailAddress;
			context.LogInfo(operation, direction, format, args);
		}

		/// <summary>
		/// Sets <paramref name="syncEntity"/> entity unique identifier to <paramref name="emailMessage"/> extended property 
		/// and saves it.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="syncEntity"><see cref="SyncEntity"/> instance.</param>
		protected void SetMessageSynchronizedInExchange(SyncContext context, SyncEntity syncEntity) {
			SetMessageSynchronizedInExchange(context, EmailMessage, syncEntity.Entity);
		}

		/// <summary>
		/// Sets <paramref name="activity"/> unique identifier to <paramref name="emailMessage"/> extended property 
		/// and saves it.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="activity"><see cref="Entity"/> instance.</param>
		protected virtual void SetMessageSynchronizedInExchange(SyncContext context,
				Exchange.EmailMessage emailMessage, Entity activity) {
			if (!context.UserConnection.GetIsFeatureEnabled("SetEmailSynchronizedInExchange")) {
				return;
			}
			try {
				emailMessage.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, activity.PrimaryColumnValue);
				emailMessage.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite);
			} catch (Exception) {
				context.LogInfo(SyncAction.Update, SyncDirection.Download, "LocalId property not set for {0} email",
					activity.PrimaryDisplayColumnValue);
			}
		}

		/// <summary>
		/// Checks that current message instance valid for synchronization.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> if message valid for synchronization. Returns <c>false</c> otherwise.</returns>
		protected bool ValidateMessage(SyncContext context, LocalItem localItem) {
			if (IsDeletedProcessed("Activity", ref localItem)) {
				LogInfo(context, Action, SyncDirection.Upload, "Item {0} already deleted.", GetDisplayName());
				return false;
			}
			if (!IsMessageValid(context)) {
				LogInfo(context, Action, SyncDirection.Upload, "Item {0} not valid.", GetDisplayName());
				Action = SyncAction.None;
				return false;
			}
			if (IsMessageLockedForSync(context)) {
				LogInfo(context, Action, SyncDirection.Upload, "Item {0} locked in another process.", GetDisplayName());
				Action = SyncAction.Repeat;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Checks is current message can be saved in bpm
		/// Message cannot be saved in current sync session if it is synchronized in another process.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c> if message can be saved in current session. Returns <c>false</c> otherwise.</returns>
		protected bool IsMessageLockedForSync(SyncContext context) {
			var helper = ClassFactory.Get<EntitySynchronizerHelper>();
			LogInfo(context, Action, SyncDirection.Upload, "Try lock Item {0}.", EmailMessage.InternetMessageId);
			return !helper.CanCreateEntityInLocalStore(EmailMessage.InternetMessageId, context.UserConnection,
				"EmailSynchronization");
		}

		/// <summary>
		/// Creates current message <see cref="SyncEntity"/> instance.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><see cref="SyncEntity"/> instance.</returns>
		protected SyncEntity GetSyncEntity(SyncContext context, LocalItem localItem) {
			var schema = context.UserConnection.EntitySchemaManager.GetInstanceByName("Activity");
			string subject = ActivityUtils.FixActivityTitle(EmailMessage.Subject, context.UserConnection);
			Entity activity = GetActivityInstance(context, localItem, schema, EmailMessage, subject);
			return localItem.Entities[schema.Name].FirstOrDefault(e => e.EntityId == activity.PrimaryColumnValue);
		}

		/// <summary>
		/// Fills <paramref name="syncEntity"/> entity with values from exchange email message.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="syncEntity"><see cref="SyncEntity"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		protected void FillActivity(SyncContext context, SyncEntity syncEntity, LocalItem localItem) {
			Entity activity = syncEntity.Entity;
			var syncProvider = context.RemoteProvider as ExchangeEmailSyncProvider;
			var sender = EmailMessage.From != null
				? EmailMessage.From.ToEmailAddressString()
				: EmailMessage.Sender.ToEmailAddressString();
			activity.SetColumnValue("Sender", sender);
			var sendDate = GetMessageDateProperty(EmailMessage, Exchange.ItemSchema.DateTimeSent, context.UserConnection);
			activity.SetColumnValue("SendDate", sendDate);
			if (syncEntity != null && syncEntity.Action != SyncAction.Create) {
				if (syncProvider != null) {
					CreateEmailMessageData(activity, syncProvider, localItem, context.UserConnection);
				}
				LogInfo(context, Action, SyncDirection.Upload, "Item {0} already exists.", GetDisplayName());
				return;
			}
			string oldBody = EmailMessage.Body.Text;
			string subject = ActivityUtils.FixActivityTitle(EmailMessage.Subject, context.UserConnection);
			FillAttachmentsDetail(context, localItem, activity.PrimaryColumnValue);
			LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] setting activity {0} fields from message {1}.",
				activity.PrimaryColumnValue, GetDisplayName());
			activity.SetColumnValue("Title", subject);
			activity.SetColumnValue("TypeId", ActivityConsts.EmailTypeUId);
			activity.SetColumnValue("Body", EmailMessage.Body.Text);
			activity.SetColumnValue("IsHtmlBody", EmailMessage.Body.BodyType == Exchange.BodyType.HTML);
			if (EmailMessage.IsIncoming()) {
				activity.SetColumnValue("HeaderProperties", EmailMessage.HeaderPropertiesToString());
			}
			activity.SetColumnValue("OwnerId", GetActivityOwnerId(EmailMessage, context.UserConnection));
			activity.SetColumnValue("Recepient", EmailMessage.ToRecipients.ToEmailAddressString());
			activity.SetColumnValue("CopyRecepient", EmailMessage.CcRecipients.ToEmailAddressString());
			activity.SetColumnValue("BlindCopyRecepient", EmailMessage.BccRecipients.ToEmailAddressString());
			var receivedDate = GetMessageDateProperty(EmailMessage, Exchange.ItemSchema.DateTimeReceived, context.UserConnection);
			SetActivityTypeFields(activity, EmailMessage);
			activity.SetColumnValue("PriorityId", EmailMessage.Importance.GetActivityPriority());
			activity.SetColumnValue("ActivityCategoryId", ActivityConsts.EmailActivityCategoryId);
			activity.SetColumnValue("StatusId", ActivityConsts.CompletedStatusUId);
			activity.SetColumnValue("DueDate", receivedDate);
			activity.SetColumnValue("StartDate", receivedDate);
			activity.SetColumnValue("MailHash", ActivityUtils.GetEmailHash(context.UserConnection, sendDate, subject, oldBody,
				context.UserConnection.CurrentUser.TimeZone));
			if (syncProvider != null) {
				activity.SetColumnValue("UserEmailAddress", syncProvider.UserSettings.SenderEmailAddress);
				LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] creating email message data for activity {0}.",
					activity.PrimaryColumnValue);
				CreateEmailMessageData(activity, syncProvider, localItem, context.UserConnection);
			}
		}

		/// <summary>
		/// Sets <paramref name="activity"/> type columns, using <paramref="emailMessage"/> instance.
		/// </summary>
		/// <param name="activity"><see cref="Entity"/> instance.</param>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		protected void SetActivityTypeFields(Entity activity, Exchange.EmailMessage emailMessage) {
			activity.SetColumnValue("EmailSendStatusId", emailMessage.GetEmailStatus());
			activity.SetColumnValue("MessageTypeId", emailMessage.GetEmailMessageType());
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Sets <paramref name="activity"/> type columns, using <paramref="emailMessage"/> instance.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		public virtual void SetActivityTypeFields(Activity activity, Exchange.EmailMessage emailMessage) {
			SetActivityTypeFields((Entity)activity, emailMessage);
		}

		/// <summary>
		/// Creates or returns existing Activity instance.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="schema">Activity schema instance.</param>
		/// <param name="message">Exchange email message.</param>
		/// <param name="subject">Email message subject.</param>
		/// <returns>Activity instance.</returns>
		/// <remarks>Added reload email relations call by task with number #CRM-13946.</remarks>
		public Activity GetActivityInstance(SyncContext context, LocalItem localItem, EntitySchema schema,
			Exchange.EmailMessage message, string subject) {
			var instance = (Activity)schema.CreateEntity(context.UserConnection);
			SyncEntity instanceSync = GetActivityInstanceSync(context, message, subject, instance);
			localItem.AddOrReplace(schema.Name, instanceSync);
			return instance;
		}

		/// <summary>
		/// Creates entities for <paramref name="localItem"/> instance using external repository item instance.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <remarks>
		/// If activity already exists in local store, then only <see cref="ActivityInFolder"/> instances updated.
		/// </remarks>
		public override void FillLocalItem(SyncContext context, ref LocalItem localItem) {
			LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] started");
			if (!ValidateMessage(context, localItem)) {
				return;
			}
			LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] validation passed, synchronization started for message {0}.", GetDisplayName());
			SyncEntity syncEntity = GetSyncEntity(context, localItem);
			FillActivityInFolders(context.UserConnection, localItem, syncEntity);
			FillActivity(context, syncEntity, localItem);
			SetMessageSynchronizedInExchange(context, syncEntity);
			LogInfo(context, Action, SyncDirection.Upload, "[FillLocalItem] for activity {0} (id {1}) ended.",
				GetDisplayName(), syncEntity.EntityId);
		}

		/// <summary><see cref="IRemoteItem.FillRemoteItem"/></summary>
		/// <remarks>Emails export not supported.</remarks>
		public override void FillRemoteItem(SyncContext context, LocalItem localItem) {
		}

		#endregion

	}

	#endregion

	#region Class: ExchangeEmailMessageUtility

	/// <summary>
	/// Provides utility methods for working with the <see cref="ExchangeEmailMessage"/> instance.
	/// </summary>
	public static class ExchangeEmailMessageUtility
	{

		#region Methods: Public

		/// <summary>
		/// Returns address string for <paramref name="emailAddress"/> instance.
		/// Address string format is <c>Name <Address></c>.
		/// </summary>
		/// <param name="emailAddress"><see cref="Exchange.EmailAddress"/> instance.</param>
		/// <returns>Formated address string.</returns>
		public static string ToEmailAddressString(this Exchange.EmailAddress emailAddress) {
			return emailAddress == null ? string.Empty : string.Format("{0} <{1}>; ", emailAddress.Name,
				emailAddress.Address).TrimStart();
		}

		/// <summary>
		/// Combines <paramref name="emailAddressCollection"/> items into one string.
		/// </summary>
		/// <remarks>Item format is <c>Name <Address>;</c></remarks>
		/// <param name="emailAddressCollection"><see cref="Exchange.EmailAddressCollection"/> instance.</param>
		/// <returns>Formated addresses string.</returns>
		public static string ToEmailAddressString(this Exchange.EmailAddressCollection emailAddressCollection) {
			var str = new StringBuilder();
			foreach (var emailAddress in emailAddressCollection) {
				str.Append(emailAddress.ToEmailAddressString());
			}
			return str.ToString();
		}

		/// <summary>
		/// Combines <paramref name="EmailMessage"/> InternetMessageHeaders items into one string.
		/// </summary>
		/// <remarks>Item format is <c>Name <PropertyName> = <c>Value <Value>;</c></remarks>
		/// <param name="EmailMessage"><see cref="xchange.EmailMessage"/> instance.</param>
		/// <returns>Formated HeaderProperties string.</returns>
		public static string HeaderPropertiesToString(this Exchange.EmailMessage emailMessage) {
			var headerProperties = new StringBuilder();
			foreach (Exchange.InternetMessageHeader property in emailMessage.InternetMessageHeaders) {
				headerProperties.Append(property.Name + "=" + property.Value + "\n");
			}
			return headerProperties.ToString();
		}

		/// <summary>
		/// Returns <paramref name="emailMessage"/> state identifier.
		/// </summary>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <returns>Message state identifier.</returns>
		public static Guid GetEmailStatus(this Exchange.EmailMessage emailMessage) {
			return emailMessage.IsDraft ? ActivityConsts.NotSendEmailStatusId : ActivityConsts.SendedEmailStatusId;
		}

		/// <summary>
		/// Returns <paramref name="emailMessage"/> type identifier.
		/// </summary>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <returns>Message type identifier.</returns>
		public static Guid GetEmailMessageType(this Exchange.EmailMessage emailMessage) {
			return emailMessage.IsIncoming() ? ActivityConsts.IncomingEmailTypeId : ActivityConsts.OutgoingEmailTypeId;
		}

		/// <summary>
		/// Returns true if <paramref name="emailMessage"/> is incoming.
		/// </summary>
		/// <param name="emailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <returns>Returns <c>true</c>, if message is incoming, <c>false</c> otherwise.</returns>
		public static bool IsIncoming(this Exchange.EmailMessage emailMessage) {
			return emailMessage.InternetMessageHeaders != null;
		}

		/// <summary>
		/// Returns <see cref="DateTime"/> instance converted to user timezone.
		/// </summary>
		/// <param name="utcDateTime"><see cref="DateTime"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns><see cref="DateTime"/> instance converted to user timezone.</returns>
		public static DateTime GetUserDateTime(this DateTime utcDateTime, UserConnection userConnection) {
			return TimeZoneInfo.ConvertTimeFromUtc(utcDateTime.Kind == DateTimeKind.Utc ?
				utcDateTime : utcDateTime.ToUniversalTime(), userConnection.CurrentUser.TimeZone);
		}

		/// <summary>
		/// Returns <paramref name="importance"/> identifier.
		/// </summary>
		/// <param name="importance">Email message <see cref="Exchange.Importance"/> instance.</param>
		/// <returns>Message importance identifier.</returns>
		public static Guid GetActivityPriority(this Exchange.Importance importance) {
			switch (importance) {
				case Exchange.Importance.Normal:
					return EmailPriorityConverter.GetActivityPriority(EmailPriority.Normal);
				case Exchange.Importance.High:
					return EmailPriorityConverter.GetActivityPriority(EmailPriority.High);
				case Exchange.Importance.Low:
					return EmailPriorityConverter.GetActivityPriority(EmailPriority.Low);
				default:
					return EmailPriorityConverter.GetActivityPriority(EmailPriority.None);
			}
		}

		#endregion
	}

	#endregion

}