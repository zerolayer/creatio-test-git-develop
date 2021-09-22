namespace Terrasoft.Configuration
{
	using System;
	using System.IO;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using System.Net;
	using System.Net.Security;
	using System.Reflection;
	using System.Security.Cryptography.X509Certificates;
	using System.Text.RegularExpressions;
	using EmailContract.DTO;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using global::Common.Logging;
	using IntegrationApi.MailboxDomain.Interfaces;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Common;
	using Terrasoft.Common.Json;
	using Terrasoft.Configuration.FileUpload;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Scheduler;
	using Terrasoft.ExchangeApi.Interfaces;
	using Terrasoft.Mail.Sender;
	using Terrasoft.Social.OAuth;
	using Terrasoft.Sync;
	using Terrasoft.Sync.Exchange;

	#region Class: ExchangeUtilityImpl

	/// <summary>
	/// Contains utility methods to work with Exchange Server Objects.
	/// </summary>
	[DefaultBinding(typeof(IExchangeUtility))]
	public class ExchangeUtilityImpl: IExchangeUtility
	{

		#region Constants: Private

		private const string SyncJobGroupName = "Exchange";
		private const string UploadAttachmentJobName = "UploadAttachment";
		private const string UploadAttachmentProcessName = "UploadActivityAttachments";
		private const string ObjectNotFoundErrorMessage = "The specified object was not found in the store";
		private const string ObjectNotFoundErrorMessageRu = "Не удается найти указанный объект в хранилище";
		private const string ObjectNotFoundErrorMessageDe = "Das angegebene Objekt wurde im Informationsspeicher nicht gefunden";
		private const string AccessDeniedErrorMessage = "Access is denied. Check credentials and try again";
		private const string AccessDeniedErrorMessageRu = "Доступ запрещен. Проверьте учетные данные и повторите попытку.";
		private const string AccessDeniedErrorMessageDe = "Der Zugriff wird verweigert. Überprüfen Sie die Anmeldeinformationen, und versuchen Sie es dann erneut.";
		private const string MailBoxNotFoundErrorMessage = "No mailbox with such guid.";
		private const string MailBoxNotFoundErrorMessageRu = "Почтовый ящик с таким GUID отсутствует.";
		private const string OccurrenceNotFoundErrorMessage = "The occurrence couldn't be found";
		private readonly Dictionary<Guid, string> ContactExtendedProperty = new Dictionary<Guid, string> {
			{
				new Guid("C11FF724-AA03-4555-9952-8FA248A11C3E"), "ContactId"
			}
		};

		#endregion

		#region Constants: Public

		/// <summary>
		/// Mail sync process name.
		/// </summary>
		public const string MailSyncProcessName = "LoadExchangeEmailsProcess";

		/// <summary>
		/// Contacts sync process name.
		/// </summary>
		public const string ContactSyncProcessName = "SyncExchangeContactsProcess";

		/// <summary>
		/// Meeting and tasks sync process name.
		/// </summary>
		public const string ActivitySyncProcessName = "SyncExchangeActivitiesProcess";

		#endregion

		#region Properties: Public

		/// <summary>
		/// Exchange integration logger.
		/// </summary>
		public ILog Log { get; } = LogManager.GetLogger("Exchange");

		/// <summary>
		/// Definition of property which contains record identifier in local storage.
		/// </summary>
		public static readonly Exchange.ExtendedPropertyDefinition LocalIdProperty = new Exchange.ExtendedPropertyDefinition(
			Exchange.DefaultExtendedPropertySet.PublicStrings, "LocalId", Exchange.MapiPropertyType.String);

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates <see cref="IAppSchedulerWraper"/> implementation instance.
		/// </summary>
		/// <returns><see cref="IAppSchedulerWraper"/> implementation instance.</returns>
		private IAppSchedulerWraper GetAppSchedulerWraper() {
			return ClassFactory.Get<IAppSchedulerWraper>();
		}

		private string GetSyncJobName(UserConnection userConnection, string jobName) {
			return jobName + "_" + userConnection.CurrentUser.Id;
		}

		private string GetSyncJobName(UserConnection userConnection, string jobName,
				string senderEmailAddress, string suffix = null) {
			string result = senderEmailAddress + "_" + jobName + "_" + userConnection.CurrentUser.Id;
			return string.IsNullOrEmpty(suffix) ? result : $"{result}_{suffix}";
		}

		/// <summary>
		/// Creates exchange synchronization process job.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="syncJobName">Synchronization job name.</param>
		/// <param name="processName">Synchronization process name.</param>
		/// <param name="period">Synchronization job run perion in minutes.</param>
		/// <param name="parameters">Synchronization process parameters.</param>
		private void CreateProcessJob(UserConnection userConnection, string syncJobName, string processName,
				int period, IDictionary<string, object> parameters) {
			RemoveProcessJob(syncJobName, userConnection);
			if (period == 0) {
				parameters.Add("CreateReminding", true);
				Log.ErrorFormat("ScheduleImmediateProcessJob called: CurrentUser {0}, SyncJobName {1}",
					userConnection.CurrentUser.Name, syncJobName);
				var appSchedulerWraper = GetAppSchedulerWraper();
				appSchedulerWraper.ScheduleImmediateProcessJob(syncJobName, SyncJobGroupName, processName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, parameters);
			} else {
				Log.ErrorFormat("ScheduleMinutelyJob called: CurrentUser {0}, SyncJobName {1}",
					userConnection.CurrentUser.Name, syncJobName);
				AppScheduler.ScheduleMinutelyJob(syncJobName, SyncJobGroupName, processName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, period, parameters);
			}
		}

		/// <summary>
		/// Creates exchange synchronization class schedule job.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="syncJobName">Synchronization job name.</param>
		/// <param name="period">Time period synchronization in minutes. If equals 0 synchronization executes once.
		/// </param>
		/// <param name="parameters">Synchronization synchronization (user email address etc.).</param>
		private void CreateClassJob(UserConnection userConnection, string syncJobName, int period,
				IDictionary<string, object> parameters) {
			RemoveProcessJob(syncJobName, userConnection);
			if (period == 0) {
				parameters.Add("CreateReminding", true);
				AppScheduler.ScheduleImmediateJob<LoadExchangeEmailsExecutor>(SyncJobGroupName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, parameters);
			} else {
				AppScheduler.ScheduleMinutelyJob<LoadExchangeEmailsExecutor>(SyncJobGroupName,
					userConnection.Workspace.Name, userConnection.Workspace.Name, period, parameters);
			}
		}

		/// <summary>
		/// Removes <paramref name="syncJobName"/> job from scheduled jobs.
		/// </summary>
		/// <param name="syncJobName">Synchronization job name.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		private void RemoveProcessJob(string syncJobName, UserConnection userConnection) {
			string trace = string.Empty;
			try {
				var stackTrace = new System.Diagnostics.StackTrace(false);
				trace = stackTrace.ToString();
			} catch(Exception e) {
				Log.Error("Error on trace creation", e);
			}
			Log.ErrorFormat("RemoveJob called: CurrentUser {0}, SyncJobName {1}. Trace: {2}",
				userConnection.CurrentUser.Name, syncJobName, trace);
			var appSchedulerWraper = GetAppSchedulerWraper();
			appSchedulerWraper.RemoveJob(syncJobName, SyncJobGroupName);
		}

		private string GetLocalizableStringValue(UserConnection userConnection, string lczName) {
			return new LocalizableString(userConnection.Workspace.ResourceStorage, "ExchangeUtility",
					"LocalizableStrings." + lczName + ".Value").ToString();
		}

		private ExchangeCredentials CreateExchangeCredentials(UserConnection userConnection,
				string emailAddress, bool ignoreRights = false) {
			var mailboxEsq = new EntitySchemaQuery(userConnection.EntitySchemaManager, "MailboxSyncSettings");
			EntitySchemaQueryColumn userNameColumn = mailboxEsq.AddColumn("UserName");
			EntitySchemaQueryColumn passwordColumn = mailboxEsq.AddColumn("UserPassword");
			EntitySchemaQueryColumn isExchengeAutodiscoverColumn = mailboxEsq.AddColumn("MailServer.IsExchengeAutodiscover");
			EntitySchemaQueryColumn exchangeEmailAddressColumn = mailboxEsq.AddColumn("MailServer.ExchangeEmailAddress");
			EntitySchemaQueryColumn useOAuth = mailboxEsq.AddColumn("MailServer.OAuthApplication.ClientClassName");
			EntitySchemaQueryColumn tokenColumn = mailboxEsq.AddColumn("OAuthTokenStorage");
			IEntitySchemaQueryFilterItem senderEmailAddressFilter = mailboxEsq
				.CreateFilterWithParameters(FilterComparisonType.Equal, "SenderEmailAddress", emailAddress);
				IEntitySchemaQueryFilterItem sysAdminUnitFilter = mailboxEsq
					.CreateFilterWithParameters(FilterComparisonType.Equal, "SysAdminUnit", userConnection.CurrentUser.Id);
			mailboxEsq.Filters.Add(senderEmailAddressFilter);
			if (!ignoreRights) {
				var filterGroup = new EntitySchemaQueryFilterCollection(mailboxEsq, LogicalOperationStrict.Or);
				filterGroup.Add(sysAdminUnitFilter);
				IEntitySchemaQueryFilterItem isSharedMailFilter = mailboxEsq
					.CreateFilterWithParameters(FilterComparisonType.Equal, "IsShared", true);
				filterGroup.Add(isSharedMailFilter);
				mailboxEsq.Filters.Add(filterGroup);
			}
			mailboxEsq.UseAdminRights = !ignoreRights;
			EntityCollection mailboxEntities = mailboxEsq.GetEntityCollection(userConnection);
			if (mailboxEntities.Count == 0) {
				throw new EmailException("ErrorOnSend", GetMailboxDoesNotExistLczValue(userConnection));
			}
			Entity mailbox = mailboxEntities[0];
			var credentials = new ExchangeCredentials {
				UserName = mailbox.GetTypedColumnValue<string>(userNameColumn.Name),
				UserPassword = mailbox.GetTypedColumnValue<string>(passwordColumn.Name),
				IsAutodiscover = mailbox.GetTypedColumnValue<bool>(isExchengeAutodiscoverColumn.Name),
				ServerAddress = mailbox.GetTypedColumnValue<string>(exchangeEmailAddressColumn.Name),
				UseOAuth = mailbox.GetTypedColumnValue<string>(useOAuth.Name).IsNotNullOrEmpty() &&
					mailbox.GetTypedColumnValue<Guid>(tokenColumn.ValueQueryAlias).IsNotEmpty()
			};
			if (credentials.UseOAuth) {
				credentials.Token = GetToken(userConnection, emailAddress);
			}
			return credentials;
		}

		private ExchangeCredentials CreateExchangeCredentials(SynchronizationCredentials credentials) {
			return new ExchangeCredentials {
				UserName = credentials.UserName,
				UserPassword = credentials.Password,
				IsAutodiscover = credentials.IsAutodiscover,
				ServerAddress = credentials.ServiceUrl,
				UseOAuth = credentials.UseOAuth
			};
		}

		private ExchangeCredentials CreateExchangeCredentials(UserConnection userConnection,
				Mail.Credentials credentials) {
			var mailServerEsq = new EntitySchemaQuery(userConnection.EntitySchemaManager, "MailServer");
			EntitySchemaQueryColumn isExchengeAutodiscoverColumn = mailServerEsq.AddColumn("IsExchengeAutodiscover");
			EntitySchemaQueryColumn exchangeEmailAddressColumn = mailServerEsq.AddColumn("ExchangeEmailAddress");
			EntitySchemaQueryColumn mailServerOAuth = mailServerEsq.AddColumn("OAuthApplication.ClientClassName");
			IEntitySchemaQueryFilterItem idFilter = mailServerEsq
				.CreateFilterWithParameters(FilterComparisonType.Equal, "Id", credentials.ServerId);
			mailServerEsq.Filters.Add(idFilter);
			EntityCollection mailServerEntities = mailServerEsq.GetEntityCollection(userConnection);
			if (mailServerEntities.Count == 0) {
				throw new EmailException("ErrorOnSend", GetMailServerDoesNotExistLczValue(userConnection));
			}
			Entity mailServer = mailServerEntities[0];
			var exCredentials = new ExchangeCredentials {
				UserName = credentials.UserName,
				UserPassword = credentials.UserPassword,
				IsAutodiscover = mailServer.GetTypedColumnValue<bool>(isExchengeAutodiscoverColumn.Name),
				ServerAddress = mailServer.GetTypedColumnValue<string>(exchangeEmailAddressColumn.Name),
				UseOAuth = mailServer.GetTypedColumnValue<string>(mailServerOAuth.Name).IsNotNullOrEmpty() && credentials.UseOAuth
			};
			return exCredentials;
		}

		private string GetToken(UserConnection userConnection, string senderEmailAddress) {
			string oauthClassName = GetOAuthClientClassNameBySender(userConnection, senderEmailAddress);
			OAuthClient oauthClient = (OAuthClient)Activator.CreateInstance(Type.GetType("Terrasoft.Configuration." + oauthClassName),
				senderEmailAddress, userConnection);
			string token = oauthClient.GetAccessToken();
			return token;
		}

		private string GetOAuthClientClassNameBySender(UserConnection userConnection, string emailSender) {
			EntitySchemaQuery esq = new EntitySchemaQuery(userConnection.EntitySchemaManager, "MailboxSyncSettings");
			EntitySchemaQueryColumn column = esq.AddColumn("MailServer.OAuthApplication.ClientClassName");
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, "SenderEmailAddress", emailSender));
			Entity mailbox = esq.GetEntityCollection(userConnection).FirstOrDefault();
			if (mailbox != null) {
				return mailbox.GetTypedColumnValue<string>(column.Name);
			} else {
				return null;
			}
		}

		/// <summary>
		/// Selects not loaded email attachments from database.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="userEmailAddress">User email address.</param>
		/// <returns>Not loaded email attachments data.</returns>
		private Dictionary<string, Dictionary<Guid, JObject>> GetAttachmentsDictionary(
				UserConnection userConnection, string userEmailAddress) {
			var result = new Dictionary<string, Dictionary<Guid, JObject>>();
			Select select = null;
			bool doNotUseMetaData = userConnection.GetIsFeatureEnabled("DoNotUseMetadataForEmail");
			if (doNotUseMetaData) {
				select = new Select(userConnection)
						.Column("AF", "ExternalStorageProperties").As("ExternalStorageProperties")
						.Column("AF", "Id").As("ActivityFileId")
					.From("ActivityFile").As("AF").WithHints(new NoLockHint())
					.InnerJoin("Activity").As("A").WithHints(new NoLockHint())
						.On("A", "Id").IsEqual("AF", "ActivityId") as Select;
			} else {
				select = new Select(userConnection).Column("SSMD", "RemoteId").As("EmailMessageId")
						.Column("SSMD", "ExtraParameters").As("AttachmentId")
						.Column("SSMD", "LocalId").As("ActivityFileId")
					.From("ActivityFile").As("AF")
					.InnerJoin("SysSyncMetaData").As("SSMD").On("SSMD", "LocalId").IsEqual("AF", "Id")
				.InnerJoin("Activity").As("A").On("A", "Id").IsEqual("AF", "ActivityId") as Select;
			}
			select = select.Where("AF", "Uploaded").IsEqual(Column.Parameter(false))
				.And().OpenBlock("AF", "ErrorOnUpload").IsEqual(Column.Parameter(string.Empty))
					.Or("AF", "ErrorOnUpload").IsNull()
					.CloseBlock()
				.And("AF", "CreatedById").IsEqual(Column.Parameter(userConnection.CurrentUser.ContactId))
				.And("A", "TypeId").IsEqual(Column.Parameter(ActivityConsts.EmailTypeUId))
				.And("A", "UserEmailAddress").IsEqual(Column.Parameter(userEmailAddress)) as Select;
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
					while (reader.Read()) {
						var activityFileId = reader.GetColumnValue<Guid>("ActivityFileId");
						var extraParams = doNotUseMetaData
							? reader.GetColumnValue<string>("ExternalStorageProperties")
							: reader.GetColumnValue<string>("AttachmentId");
						var attachmentId = TryGetPropertyFromJson(extraParams, "RemoteId");
						var activityId = TryGetPropertyFromJson(extraParams, "ActivityId");
						var emailMessageId = doNotUseMetaData
							? TryGetPropertyFromJson(extraParams, "ExchangeMessageId")
							: reader.GetColumnValue<string>("EmailMessageId");
						var jData = new JObject() {
									{ "ActivityId", activityId.ToString()},
									{ "AttachmentId", attachmentId}
								};
						if (!result.ContainsKey(emailMessageId)) {
							result[emailMessageId] = new Dictionary<Guid, JObject> { { activityFileId, jData } };
						} else {
							result[emailMessageId].Add(activityFileId, jData);
						}
					}
				}
			}
			return result;
		}

		/// <summary>
		/// Saves attachment data.
		/// </summary>
		/// <param name="userConnection">User connection instance.</param>
		/// <param name="content">Attachment content.</param>
		/// <param name="activityFileId">Activity file id.</param>
		/// <param name="name">Name</param>
		private void SaveAttachmentData(UserConnection userConnection,
			byte[] content, Guid activityFileId, string name = null) {
			var fileRepository = ClassFactory.Get<FileRepository>(new ConstructorArgument("userConnection", userConnection));
			var fileEntityInfo = new FileEntityUploadInfo("ActivityFile", activityFileId, name) {
				Content = new MemoryStream(content),
				TotalFileLength = content.Length
			};
			fileRepository.UploadFile(fileEntityInfo, false);
			var update = new Update(userConnection, "ActivityFile")
							.Set("Uploaded", Column.Parameter(true))
						.Where("Id").IsEqual(Column.Parameter(activityFileId));
			update.Execute();
		}

		private void RaiseAttachmentError(UserConnection userConnection, string errorMessage,
				Guid activityFileId) {
			var update = new Update(userConnection, "ActivityFile")
							.Set("ErrorOnUpload", Column.Parameter(errorMessage))
						.Where("Id").IsEqual(Column.Parameter(activityFileId));
			update.Execute();
		}

		private void CreateNewAttach(UserConnection userConnection, Exchange.Attachment attachment,
				Guid activityId, IExchangeAttachmentUtilities attachmentHelper) {
			Exchange.Attachment emailAttachment = attachment as Exchange.FileAttachment;
			if (emailAttachment == null) {
				emailAttachment = attachment as Exchange.ItemAttachment;
			}
			if (emailAttachment != null) {
				Guid fileId = Guid.NewGuid();
				var insert = new Insert(userConnection)
					.Set("ErrorOnUpload", Column.Parameter("Not existed attach"))
					.Set("Name", Column.Parameter(emailAttachment.Name))
					.Set("Id", Column.Parameter(fileId))
					.Set("ActivityId", Column.Parameter(activityId))
					.Into("ActivityFile");
				insert.Execute();
				try {
					attachmentHelper.Load(emailAttachment);
					SaveAttachmentData(userConnection, attachmentHelper.GetContent(emailAttachment), fileId, emailAttachment.Name);
				} catch (Exchange.ServiceResponseException e) {
					RaiseAttachmentError(userConnection, e.Message, fileId);
				}
			}
		}

		/// <summary>
		/// Uploads attachment body for not loaded <see cref="ActivityFile"/>.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="attachment">Email attachment.</param>
		/// <param name="attachmentHelper"><see cref="IExchangeAttachmentUtilities"/> utilities object.</param>
		/// <param name="message">Email message.</param>
		private bool UploadAttachmentData(UserConnection userConnection, KeyValuePair<Guid, JObject> attachment,
				IExchangeAttachmentUtilities attachmentHelper, Exchange.EmailMessage message) {
			var jData = attachment.Value;
			string attachmentId = jData.Value<string>("AttachmentId");
			ILog log = LogManager.GetLogger("Exchange");
			IEnumerable<Exchange.Attachment> attachments = attachmentHelper.GetAttachmentsById(message, attachmentId);
			try {
				if (!attachments.Any()) {
					RaiseAttachmentError(userConnection, "Attachment does not exist", attachment.Key); 
				} else {
					Exchange.Attachment firstAttachment = attachments.First();
					Exchange.Attachment emailAttachment = firstAttachment as Exchange.FileAttachment;
					if (emailAttachment == null) {
						emailAttachment = firstAttachment as Exchange.ItemAttachment;
					}
					if (emailAttachment != null) {
						attachmentHelper.Load(emailAttachment);
						SaveAttachmentData(userConnection, attachmentHelper.GetContent(emailAttachment), attachment.Key, emailAttachment.Name);
						return true;
					} else {
						RaiseAttachmentError(userConnection, "Attachment does not exist", attachment.Key);
					}
				}
			} catch (Exception exception) {
				log.ErrorFormat("AttachmentId {0}, currentUser {1}", exception, attachment.Key, userConnection.CurrentUser.Id);
				RaiseAttachmentError(userConnection, exception.Message, attachment.Key);
			}
			return false;
		}

		/// <summary>
		/// Checks is <paramref name="message"/> is not processed errors message.
		/// </summary>
		/// <param name="message">Exception message.</param>
		/// <returns><c>True</c> if message not processed. Returns <c>false</c> otherwise.</returns>
		private bool IsNotProcessedErrorMessages(string message) {
			return new List<string> { ObjectNotFoundErrorMessage, ObjectNotFoundErrorMessageRu, ObjectNotFoundErrorMessageDe,
					OccurrenceNotFoundErrorMessage }.All(m => !message.Contains(m));
		}

		private bool IsItemProcessedErrorMessage(string message) {
			return new List<string> { AccessDeniedErrorMessage, AccessDeniedErrorMessageRu, AccessDeniedErrorMessageDe,
						MailBoxNotFoundErrorMessage, MailBoxNotFoundErrorMessageRu }.Any(m => message.Contains(m));
		}

		private void SetSecurityProtocolOptions(UserConnection userConnection) {
			if (userConnection.GetIsFeatureEnabled("ExtendedSslTlsList")) {
				ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12 |
					SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
			}
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
		}

		private void ResetSecurityProtocolOptions() {
			ServicePointManager.ServerCertificateValidationCallback -= ValidateRemoteCertificate;
		}

		#endregion

		#region Methods: Internal

		/// <summary>
		/// Tests <see cref="Exchange.ExchangeService"/> connection.
		/// </summary>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		internal void TestConnection(Exchange.ExchangeService service, UserConnection userConnection = null,
				string senderEmailAddress = "", bool stopOnFirstError = false) {
			SynchronizationErrorHelper helper = SynchronizationErrorHelper.GetInstance(userConnection);
			try {
				var id = new Exchange.FolderId(Exchange.WellKnownFolderName.MsgFolderRoot, senderEmailAddress);
				service.FindFolders(id, new Exchange.FolderView(1));
				helper.CleanUpSynchronizationError(senderEmailAddress);
			} catch (Exception ex) {
				helper.ProcessSynchronizationError(senderEmailAddress, ex, stopOnFirstError);
				throw;
			}
		}

		/// <summary>
		/// Creates <see cref="Exchange.ExchangeService"/> instance.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="credentials"><see cref="ExchangeCredentials"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		/// <returns><see cref="Exchange.ExchangeService"/> instance.</returns>
		internal Exchange.ExchangeService CreateExchangeService(UserConnection userConnection, ExchangeCredentials credentials,
				string senderEmailAddress, bool stopOnFirstError = false) {
			SetSecurityProtocolOptions(userConnection);
			Exchange.ExchangeService exchangeService;
			try {
				exchangeService = new Exchange.ExchangeService(Exchange.ExchangeVersion.Exchange2010_SP1, TimeZoneInfo.Utc);
				if (userConnection.GetIsFeatureEnabled("TraceEws")) {
					exchangeService.TraceListener = new TraceListenerInstance(Log, senderEmailAddress);
					exchangeService.TraceFlags = Exchange.TraceFlags.All;
					exchangeService.TraceEnabled = true;
				}
				if (credentials.IsAutodiscover) {
					exchangeService.AutodiscoverUrl(senderEmailAddress, delegate (string url) {
						var regexString = Terrasoft.Core.Configuration.SysSettings.GetValue<string>(userConnection,
							"ExchangeAutodiscoverRedirectUrlRegex", ".*");
						return Regex.IsMatch(url, regexString, RegexOptions.IgnoreCase);
					});
				} else {
					exchangeService.Url = new Uri(string.Format("https://{0}/EWS/Exchange.asmx", credentials.ServerAddress));
				}
				if (credentials.UseOAuth) {
					string token = GetToken(userConnection, senderEmailAddress);
					exchangeService.Credentials = new Exchange.OAuthCredentials(token);
				} else {
					exchangeService.Credentials = new Exchange.WebCredentials(credentials.UserName, credentials.UserPassword);
				}
				TestConnection(exchangeService, userConnection, senderEmailAddress, stopOnFirstError);
			} catch (Exception e) {
				Log.Error(e.Message);
				throw;
			} finally {
				ResetSecurityProtocolOptions();
			}
			return exchangeService;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Checks if contact metadata exists by <see cref="contactId"/>.
		/// </summary>
		/// <param name="contactId">Contact uniqueidentifier.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>True if contact metadata exists. </returns>
		public bool IsContactMetadataExist(Guid contactId, UserConnection userConnection) {
			Select selectContactMetadata =
				new Select(userConnection).Top(1)
					.Column("Id")
				.From("SysSyncMetaData")
				.Where("LocalId").IsEqual(Column.Parameter(contactId)) as Select;
			return selectContactMetadata.ExecuteScalar<Guid>().IsNotEmpty();
		}

		/// <summary>
		/// Returns <see cref="Exchange.Item"/> importance.
		/// </summary>
		/// <param name="activityPriorityId"><see cref="ActivityPriority"/> instance id.</param>
		/// <returns><see cref="Exchange.Importance"/> instance.</returns>
		public int GetExchangeImportance(Guid activityPriorityId) {
			if (activityPriorityId == ExchangeConsts.ActivityHighPriorityId) {
				return (int)Exchange.Importance.High;
			} else if (activityPriorityId == ExchangeConsts.ActivityNormalPriorityId) {
				return (int)Exchange.Importance.Normal;
			}
			return (int)Exchange.Importance.Low;
		}

		/// <summary>
		/// Returns <see cref="Exchange.Task"/> status.
		/// </summary>
		/// <param name="activityStatusId"><see cref="ActivityStatus"/> instance id.</param>
		/// <returns><see cref="Exchange.TaskStatus"/> instance.</returns>
		public int GetExchangeTaskStatus(Guid activityStatusId) {
			if (activityStatusId == ActivityConsts.InProgressUId) {
				return (int)Exchange.TaskStatus.InProgress;
			} else if (activityStatusId == ActivityConsts.CompletedStatusUId) {
				return (int)Exchange.TaskStatus.Completed;
			}
			return (int)Exchange.TaskStatus.NotStarted;
		}

		/// <summary>
		/// Returns <see cref="key"/> value from external repository values dictionary <paramref name="dict"/>.
		/// If key value not reachable, returns default value.
		/// </summary>
		/// <typeparam name="TKey">Dictionary key value type.</typeparam>
		/// <typeparam name="TEntry">Dictionary entry value type.</typeparam>
		/// <typeparam name="TOutputType">Return value type.</typeparam>
		/// <param name="dict">External repository detail values dictionary.</param>
		/// <param name="key">Key value <paramref name="dict"/>.</param>
		/// <returns><see cref="key"/> value from external repository values dictionary <paramref name="dict"/>
		/// if value reachable, dafault value otherwise.</returns>
		public TOutputType SafeGetValue<TKey, TEntry, TOutputType>(Exchange.DictionaryProperty<TKey,
				TEntry> dict, TKey key)
				where TEntry : Exchange.DictionaryEntryProperty<TKey> {
			TOutputType result = default(TOutputType);
			Type type = dict.GetType();
			MethodInfo getMethod = type.GetMethod("TryGetValue");
			var args = new object[] { key, result };
			getMethod.Invoke(dict, args);
			return (TOutputType)args[1];
		}

		/// <summary>
		/// Delete synchronizer from activitySynchronizer.
		/// </summary>
		/// <param name="userConnection">A instance of the current user connection.</param>
		/// <returns></returns>
		public void DeleteEmptyActivityFromActivitySynchronizer(UserConnection userConnection,
				Guid activityType) {
			var delete = new Delete(userConnection)
						.From("ActivitySynchronizer")
						.Where("ActivityId").Not().Exists(
							new Select(userConnection)
								.Column("Id")
							.From("Activity")
							.Where("ActivitySynchronizer", "ActivityId").IsEqual("Activity", "Id")
								.And("Activity", "TypeId").IsEqual(Column.Parameter(activityType)));
			using (var dbExecutor = userConnection.EnsureDBConnection()) {
				delete.Execute(dbExecutor);
			}
		}

		/// Verifies that the current user is the synchronizer of the activity by <paramref name="remoteUId"/>.
		/// <param name="userConnection">User connection.</param>
		/// <param name="remoteUId">The unique identifier of the activity on remote storage.</param>
		/// <returns>Flag of conformity user and the synchronizer.
		/// </returns>
		public bool CheckSynchronizer(UserConnection userConnection, string remoteUId) {
			Guid currentContactId = userConnection.CurrentUser.ContactId;
			Guid synchronizerId = GetSynchronizerByRemoteUId(userConnection, remoteUId);
			return (!synchronizerId.IsEmpty()) && (synchronizerId != currentContactId);
		}

		/// <summary>
		/// Set synchronizer in activitySynchronizer.
		/// </summary>
		/// <param name="userConnection">A instance of the current user connection.</param>
		/// <param name="remoteUId">The unique identifier of the activity on remote storage.</param>
		/// <param name="activityId">The unique identifier of the activity.</param>
		/// <param name="organizerId">The unique identifier of the organizer.</param>
		/// <returns></returns>
		public bool SetSynchronizer(UserConnection userConnection, Guid organizerId, Guid activityId,
				string remoteUId) {
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				try {
					dbExecutor.StartTransaction();
					Guid SynchronizerId = GetSynchronizerByRemoteUId(userConnection, remoteUId);
					if (SynchronizerId == organizerId) {
						dbExecutor.CommitTransaction();
						return true;
					}
					if (SynchronizerId.IsEmpty()) {
						var insert =
							new Insert(userConnection).Into("ActivitySynchronizer")
								.Set("RemoteUId", Column.Parameter(remoteUId))
								.Set("CreatedById", Column.Parameter(organizerId))
								.Set("ModifiedById", Column.Parameter(organizerId))
								.Set("ActivityId", Column.Parameter(activityId));
						insert.Execute(dbExecutor);
						dbExecutor.CommitTransaction();
						return true;
					}
					dbExecutor.RollbackTransaction();
					return false;
				} catch (Exception ex) {
					dbExecutor.RollbackTransaction();
					Log.ErrorFormat("[ExchangeUtility.SetSynchronizer]: Error on setting synchronizer: {0}", ex.Message);
					return false;
				}
			}
		}

		/// <summary>
		/// Get synchronizer from activitySynchronizer.
		/// </summary>
		/// <param name="userConnection">A instance of the current user connection.</param>
		/// <param name="remoteUId">The unique identifier of the activity on remote storage.</param>
		/// <returns>The unique identifier of the synchronizer.</returns>
		public Guid GetSynchronizerByRemoteUId(UserConnection userConnection, string remoteUId) {
			Select SelectActivity =
				new Select(userConnection).Column("CreatedById")
					.From("ActivitySynchronizer")
						.Where("RemoteUId").IsEqual(Column.Parameter(remoteUId)) as Select;
			return SelectActivity.ExecuteScalar<Guid>();
		}

		/// <summary>
		/// Set <paramref name="extraParameters"/> into <paramref name="syncEntity"/>.
		/// </summary>
		/// <param name="syncEntity">The object of activity synchronization.</param>
		/// <param name="extraParameters">The json object with extra parameters.</param>
		public void SetActivityExtraParameters(SyncEntity syncEntity, JObject extraParameters) {
			syncEntity.ExtraParameters = Json.Serialize(extraParameters);
		}

		/// <summary>
		/// Returns exchange store item instance with <paramref name="id"/>.
		/// Returns <c>null</c> if item is not reachable.
		/// </summary>
		/// <typeparam name="T">Requested <see cref="Exchange.Item"/> type.</typeparam>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="id"><see cref="Exchange.ItemId"/> instance.</param>
		/// <returns><see cref="Exchange.Item"/> instance if item exists. Otherwise return <c>null</c>.</returns>
		public T SafeBindItem<T>(Exchange.ExchangeService service, Exchange.ItemId id)
				where T : Exchange.Item {
			T value = null;
			try {
				value = Exchange.Item.Bind(service, id) as T;
			} catch (Exchange.ServiceResponseException exception) {
				string message = exception.Message;
				if (IsItemProcessedErrorMessage(message)) {
					Log.ErrorFormat("[ExchangeUtility.SafeBindItem]: Error while loading item with Id: {0}", exception, id);
				} else if (IsNotProcessedErrorMessages(message)) {
					throw;
				}
			}
			return value;
		}

		/// <summary>
		/// Returns exchange store item instance with <paramref name="id"/>.
		/// Returns <c>null</c> if item is not reachable.
		/// </summary>
		/// <typeparam name="T">Requested <see cref="Exchange.Item"/> type.</typeparam>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="id"><see cref="Exchange.ItemId"/> instance.</param>
		/// <param name="propertySet"><see cref="Exchange.PropertySet"/> instance defines properties to load.</param>
		/// <returns><see cref="Exchange.Item"/> instance if item exists. Otherwise return <c>null</c>.</returns>
		public T SafeBindItem<T>(Exchange.ExchangeService service, Exchange.ItemId id, Exchange.PropertySet propertySet)
			where T : Exchange.Item {
			T value = null;
			try {
				value = Exchange.Item.Bind(service, id, propertySet) as T;
			} catch (Exchange.ServiceResponseException exception) {
				string message = exception.Message;
				if (IsItemProcessedErrorMessage(message)) {
					Log.ErrorFormat("[ExchangeUtility.SafeBindItem]: Error while loading item with Id: {0}", exception, id);
				} else if (IsNotProcessedErrorMessages(message)) {
					throw;
				}
			}
			return value;
		}


		/// <summary>
		/// In the directory that corresponds to the specified name <paramref name = "folderName"/>,
		/// searches for all subdirectories that match the specified filter <paramref name="filter"/>,
		/// and returns a hierarchical collection <see cref="Terrasoft.Configuration.ExchangeMailFolder"/>.
		/// </summary>
		/// <param name="service">Binding to the Exchange Web service.</param>
		/// <param name="folderName">Common folder name used in user mailboxes.</param>
		/// <param name="filter">Search filter.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="selectedFolders">Enumerator of the unique identifiers of the selected directories. The default value is <c>null</c>.</param>
		/// <returns>Collection objects of <see cref="Terrasoft.Configuration.ExchangeMailFolder"/>.
		/// </returns>
		public List<ExchangeMailFolder> GetHierarhicalFolderList(Exchange.ExchangeService service,
				Exchange.WellKnownFolderName folderName, Exchange.SearchFilter filter,
				string senderEmailAddress, IEnumerable<string> selectedFolders = null) {
			var id = new Exchange.FolderId(folderName, senderEmailAddress);
			Exchange.Folder folder = Exchange.Folder.Bind(service, id);
			return GetHierarhicalFolderList(folder, filter, selectedFolders);
		}

		/// <summary>
		/// In the specified directory <paramref name="rootFolder"/> searches for all subdirectories, that match the specified filter <paramref name="filter"/>
		/// and returns a hierarchical collection <see cref="Terrasoft.Configuration.ExchangeMailFolder"/>.
		/// </summary>
		/// <param name="rootFolder">The directory from which the recursive search is performed.</param>
		/// <param name="filter">Search filter.</param>
		/// <param name="selectedFolders">Enumerator of the unique identifiers of the selected directories. The default value is <c>null</c>.</param>
		/// <returns>Collection objects of <see cref="ExchangeMailFolder"/>.</returns>
		public List<ExchangeMailFolder> GetHierarhicalFolderList(Exchange.Folder rootFolder,
				Exchange.SearchFilter filter,  IEnumerable<string> selectedFolders = null) {
			var result = new List<ExchangeMailFolder>();
			var folders = new List<Exchange.Folder>();
			folders.GetAllFoldersByFilter(rootFolder, filter);
			if (!folders.Any()) {
				return null;
			}
			var rootId = Guid.NewGuid().ToString();
			result.Add(new ExchangeMailFolder {
				Id = rootId,
				Name = rootFolder.DisplayName,
				ParentId = string.Empty,
				Path = string.Empty,
				Selected = false
			});
			foreach (Exchange.Folder folder in folders) {
				ExchangeMailFolder parentFolder =
					result.FirstOrDefault(item => item.Path.Equals(folder.ParentFolderId.UniqueId));
				string parentId = parentFolder.Id ?? rootId;
				result.Add(new ExchangeMailFolder {
					Id = Guid.NewGuid().ToString(),
					Name = folder.DisplayName,
					ParentId = parentId,
					Path = folder.Id.UniqueId,
					Selected = selectedFolders != null && selectedFolders.Contains(folder.Id.UniqueId)
				});
			}
			return result;
		}

		/// <summary>
		/// Creates <see cref="Exchange.ExchangeService"/> instance for <paramref name="senderEmailAddress"/> settings.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		/// <returns><see cref="Exchange.ExchangeService"/> instance.</returns>
		public Exchange.ExchangeService CreateExchangeService(UserConnection userConnection,
				string senderEmailAddress, bool stopOnFirstError = false, bool ignoreRights = false) {
			ExchangeCredentials exchangeCredentials = CreateExchangeCredentials(userConnection, senderEmailAddress, ignoreRights);
			string logTemplate = "ExchangeClient.CreateExchangeService(UserConnection , senderEmailAddress) -> " +
							" Created ExchangeCredentials (ServerAddress:'{0}', UserName:'{1}', senderEmailAddress:'{2}')" +
							" and create ExchangeService'.\r\n";
			Log.InfoFormat(logTemplate, exchangeCredentials.ServerAddress,
					((Terrasoft.Mail.Credentials)(exchangeCredentials)).UserName, senderEmailAddress);
			Exchange.ExchangeService exchangeService = CreateExchangeService(userConnection, exchangeCredentials, senderEmailAddress, stopOnFirstError);
			return exchangeService;
		}

		/// <summary>
		/// Creates <see cref="Exchange.ExchangeService"/> instance using <paramref name="credentials"/> credentials.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="credentials">Exchane service credentials.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		/// <returns><see cref="Exchange.ExchangeService"/> instance.</returns>
		public Exchange.ExchangeService CreateExchangeService(UserConnection userConnection,
				Mail.Credentials credentials, string senderEmailAddress, bool stopOnFirstError = false) {
			ExchangeCredentials exchangeCredentials = CreateExchangeCredentials(userConnection, credentials);
			string logTemplate = "ExchangeClient.CreateExchangeService(UserConnection , Mail.Credentials, senderEmailAddress) -> " +
							" Created ExchangeCredentials (ServerAddress:'{0}', UserName:'{1}', senderEmailAddress:'{2}')" +
							" and create ExchangeService'.\r\n";
			Log.InfoFormat(logTemplate, exchangeCredentials.ServerAddress,
					((Terrasoft.Mail.Credentials)(exchangeCredentials)).UserName, senderEmailAddress);
			Exchange.ExchangeService exchangeService = CreateExchangeService(userConnection, exchangeCredentials, senderEmailAddress, stopOnFirstError);
			return exchangeService;
		}

		/// <summary>
		/// Creates <see cref="Exchange.ExchangeService"/> instance using <paramref name="credentials"/> credentials.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="credentials">Exchange service credentials.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		/// <returns><see cref="Exchange.ExchangeService"/> instance.</returns>
		public Exchange.ExchangeService CreateExchangeService(UserConnection userConnection,
				SynchronizationCredentials credentials, string senderEmailAddress, bool stopOnFirstError = false) {
			ExchangeCredentials exchangeCredentials = CreateExchangeCredentials(credentials);
			string logTemplate = "ExchangeClient.CreateExchangeService(UserConnection , Mail.Credentials, senderEmailAddress) -> " +
							" Created ExchangeCredentials (ServerAddress:'{0}', UserName:'{1}', senderEmailAddress:'{2}')" +
							" and create ExchangeService.\r\n";
			Log.InfoFormat(logTemplate, exchangeCredentials.ServerAddress,
					((Terrasoft.Mail.Credentials)(exchangeCredentials)).UserName, senderEmailAddress);
			Exchange.ExchangeService exchangeService = CreateExchangeService(userConnection, exchangeCredentials, senderEmailAddress, stopOnFirstError);
			return exchangeService;
		}

		/// <summary>
		/// Returns localized "Mailbox does not exists" error message.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Locaized error message.</returns>
		public string GetMailboxDoesNotExistLczValue(UserConnection userConnection) {
			return GetLocalizableStringValue(userConnection, "MailboxDoesNotExist");
		}

		/// <summary>
		/// Returns localized "Mail provider does not exists" error message.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Locaized error message.</returns>
		public string GetMailServerDoesNotExistLczValue(UserConnection userConnection) {
			return GetLocalizableStringValue(userConnection, "MailServerDoesNotExist");
		}

		/// <summary>
		/// Returns localized "Process synchronization failed" error message.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="processName">Exchange synchronization process name.</param>
		/// <returns>Locaized error message.</returns>
		public string GetProcessErrorExecutionLczValue(UserConnection userConnection, string processName) {
			return GetLocalizableStringValue(userConnection, processName + "ErrorExecution");
		}

		/// <summary>
		/// Returns localized "The current user is not the owner of this mailbox" error message.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Locaized error message.</returns>
		public string GetUserIsNotMailboxOwnerLczValue(UserConnection userConnection) {
			return GetLocalizableStringValue(userConnection, "UserIsNotMailboxOwner");
		}

		/// <summary>
		/// Returns localized caption "Private meeting".
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Localized caption for private meeting.</returns>
		public string GetPrivateAppointmentTitleLczValue(UserConnection userConnection) {
			return ActivityUtils.GetLczPrivateMeeting(userConnection);
		}

		/// <summary>
		/// Creates load attachments from Exchange scheduler job.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="userEmailAddress">Sender email address.</param>
		public void CreateUploadAttachmentJob(UserConnection userConnection, string userEmailAddress) {
			string syncJobName = GetSyncJobName(userConnection, UploadAttachmentJobName, userEmailAddress);
			RemoveProcessJob(syncJobName, userConnection);
			Dictionary<string, object> parameters = new Dictionary<string, object> {
				{ "UserEmailAddress", userEmailAddress }
			};
			var schedulerWrapper = GetAppSchedulerWraper();
			if (userConnection.GetIsFeatureEnabled("UseClassSynchronizer")) {
				schedulerWrapper.ScheduleImmediateJob<UploadAttachmentsDataExecutor>(SyncJobGroupName, userConnection.Workspace.Name,
					userConnection.CurrentUser.Name, parameters);
			} else {
				Log.InfoFormat("ScheduleImmediateProcessJob called: CurrentUser {0}, SyncJobName {1}",
					userConnection.CurrentUser.Name, syncJobName);
				schedulerWrapper.ScheduleImmediateProcessJob(syncJobName, SyncJobGroupName, UploadAttachmentProcessName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, parameters);
			}
		}

		/// <summary>
		/// Creates exchange synchronization process schedule job.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="periodInMinutes">Job interval.</param>
		/// <param name="processName">Synchronization process name.</param>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <remarks>
		/// If <paramref name="periodInMinutes"/> is 0, imediate job will be created.
		/// </remarks>
		public void CreateSyncJob(UserConnection userConnection, int periodInMinutes,
				string processName, IDictionary<string, object> parameters = null) {
			string suffix = string.Empty;
			if (periodInMinutes < 0) {
				throw new ArgumentOutOfRangeException("periodInMinutes");
			}
			if (periodInMinutes == 0) {
				suffix = "ImmediateProcessJob";
			}
			string syncJobName;
			if (parameters != null && parameters.ContainsKey("SenderEmailAddress")) {
				syncJobName = GetSyncJobName(userConnection, processName, parameters["SenderEmailAddress"].ToString(), suffix);
			} else {
				syncJobName = GetSyncJobName(userConnection, processName);
			}
			if (userConnection.GetIsFeatureEnabled("UseClassSynchronizer")) {
				CreateClassJob(userConnection, syncJobName, periodInMinutes, parameters);
			} else {
				CreateProcessJob(userConnection, syncJobName, processName, periodInMinutes, parameters);
			}
		}

		/// <summary>
		/// Deletes exchange synchronization process schedule job.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email adress.</param>
		/// <param name="processName">Synchronization process name.</param>
		public void RemoveSyncJob(UserConnection userConnection, string senderEmailAddress,
				string processName) {
			string syncJobName = GetSyncJobName(userConnection, processName, senderEmailAddress);
			RemoveProcessJob(syncJobName, userConnection);
		}

		/// <summary>
		/// Check if sync job exists.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="processName">Synchronization process name.</param>
		/// <returns><c>True</c> if job exists. Otherwise returns <c>false</c>.</returns>
		public bool DoesSyncJobExists(UserConnection userConnection, string senderEmailAddress,
				string processName) {
			string syncJobName = GetSyncJobName(userConnection, processName, senderEmailAddress);
			var appSchedulerWraper = GetAppSchedulerWraper();
			return appSchedulerWraper.DoesJobExist(syncJobName, SyncJobGroupName);
		}

		/// <summary>
		/// Deletes jobs for mailbox synchronization.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email adress.</param>
		public void RemoveAllSyncJob(UserConnection userConnection, string senderEmailAddress, Guid serverTypeId) {
			if (serverTypeId != Guid.Empty && serverTypeId == ExchangeConsts.ExchangeMailServerTypeId) {
				RemoveSyncJob(userConnection, senderEmailAddress, MailSyncProcessName);
				RemoveSyncJob(userConnection, senderEmailAddress, ContactSyncProcessName);
				RemoveSyncJob(userConnection, senderEmailAddress, ActivitySyncProcessName);
			}
		}

		/// <summary>
		/// Creates Exchange contact extended property instance.
		/// </summary>
		/// <returns><see cref="Exchange.ExtendedPropertyDefinition"/> instance.</returns>
		public Exchange.ExtendedPropertyDefinition GetContactExtendedPropertyDefinition() {
			return new Exchange.ExtendedPropertyDefinition(ContactExtendedProperty.First().Key,
				ContactExtendedProperty.First().Value, Exchange.MapiPropertyType.String);
		}

		/// <summary>
		/// Returns contact extended property name.
		/// </summary>
		/// <returns>Exchange contact extended property name string.</returns>
		public string GetContactExtendedPropertyName() {
			return ContactExtendedProperty.First().Value;
		}

		/// <summary>
		/// Uploads attachments body for not loaded <see cref="ActivityFile"/>.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="userEmailAddress">User email address.</param>
		/// <remarks>
		/// <see cref="ExchangeAttachmentHelper"/> used by default.
		/// </remarks>
		public void UploadAttachmentsData(UserConnection userConnection, string userEmailAddress) {
			ConstructorArgument argument = new ConstructorArgument("userConnection", userConnection);
			IExchangeAttachmentUtilities attachmentHelper = ClassFactory.Get<IExchangeAttachmentUtilities>(userConnection.GetType().Name, argument);
			Dictionary<string, Dictionary<Guid, JObject>> attachmentsDictionary =
				GetAttachmentsDictionary(userConnection, userEmailAddress);
			Exchange.ExchangeService service;
			try {
				service = attachmentHelper.CreateExchangeService(userConnection, userEmailAddress);
			} catch (Exchange.ServiceResponseException) {
				return;
			}
			SetSecurityProtocolOptions(userConnection);
			Exchange.PropertySet properties = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly, Exchange.ItemSchema.Attachments);
			foreach (KeyValuePair<string, Dictionary<Guid, JObject>> email in attachmentsDictionary) {
				if (email.Key.IsNullOrEmpty()) {
					email.Value.ForEach(att => RaiseAttachmentError(userConnection, "EmailMessage Id does not set in ExternalStorageProperties",
						att.Key));
					continue;
				}
				Exchange.EmailMessage message = attachmentHelper.SafeBindItem(service, new Exchange.ItemId(email.Key), properties);
				if (message == null) {
					email.Value.ForEach(att => RaiseAttachmentError(userConnection, "EmailMessage does not exist",
						att.Key));
					continue;
				}
				try {
					attachmentHelper.GetAttachments(message);
				} catch (Exchange.ServiceResponseException) {
					email.Value.ForEach(att => RaiseAttachmentError(userConnection, "Message attachments does not exist",
						att.Key));
					continue;
				}
				Guid.TryParse(email.Value.First().Value.Value<string>("ActivityId"), out Guid activityId);
				var processedAttachs = new List<Exchange.Attachment>();
				foreach (KeyValuePair<Guid, JObject> attachment in email.Value) {
					if (UploadAttachmentData(userConnection, attachment, attachmentHelper, message)) {
						string attachmentId = attachment.Value.Value<string>("AttachmentId");
						processedAttachs.Add(attachmentHelper.GetAttachmentsById(message, attachmentId).First());
					}
				}
				foreach (Exchange.Attachment attachment in attachmentHelper.GetAttachments(message)) {
					if (!processedAttachs.Contains(attachment) && activityId.IsNotEmpty()) {
						CreateNewAttach(userConnection, attachment, activityId, attachmentHelper);
					}
				}
			}
			ResetSecurityProtocolOptions();
		}

		/// <summary>
		/// Starts exchange items synchronization process.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">User email address.</param>
		/// <param name="exchangeSyncProviderFunc">Delegate function that returns
		/// an instance of the Exchange synchronization provider.</param>
		/// <param name="resultMessage">Synchronization result message.</param>
		/// <param name="localChangesCount">Local storage changes count.</param>
		/// <param name="remoteChangesCount">Remote storage changes count.</param>
		/// <param name="syncProcessName">Synchronization process name.</param>
		public void SyncExchangeItems(UserConnection userConnection, string senderEmailAddress,
				Func<BaseExchangeSyncProvider> exchangeSyncProviderFunc, out string resultMessage,
				out int localChangesCount, out int remoteChangesCount, string syncProcessName = null) {
			resultMessage = null;
			localChangesCount = 0;
			remoteChangesCount = 0;
			try {
				var localProvider = new LocalProvider(userConnection);
				var remoteProvider = exchangeSyncProviderFunc();
				var synContext = new SyncContext {
					UserConnection = userConnection,
					LocalProvider = localProvider,
					RemoteProvider = remoteProvider
				};
				var synAgent = new SyncAgent(synContext);
				synAgent.LocalItemAppliedInRemoteStore += remoteProvider.OnLocalItemAppliedInRemoteStore;
				SetSecurityProtocolOptions(userConnection);
				synAgent.Synchronize();
				synAgent.LocalItemAppliedInRemoteStore -= remoteProvider.OnLocalItemAppliedInRemoteStore;
				ResetSecurityProtocolOptions();
				localChangesCount = remoteProvider.LocalChangesCount;
				remoteChangesCount = remoteProvider.RemoteChangesCount;
			} catch (Exception e) {
				Log.Error($"Error in {senderEmailAddress} synchronization", e);
				resultMessage = e.Message;
			}
		}

		/// <summary>
		/// Extracts <paramref name="propertyName"/> value from serialized <paramref name="extraParameters"/> json string.
		/// </summary>
		/// <param name="extraParameters">Serialized json string.</param>
		/// <param name="propertyName">Property value.</param>
		/// <returns>Property value from json.</returns>
		/// <remarks>
		/// Returns <paramref name="extraParameters"/> string if string not valid json
		/// or does not contains <paramref name="propertyName"/> property.
		/// </remarks>
		public string TryGetPropertyFromJson(string extraParameters, string propertyName) {
			try {
				JObject extraParametersObj = JObject.Parse(extraParameters);
				var propertyValue = (string)extraParametersObj[propertyName];
				return string.IsNullOrEmpty(propertyValue) ? extraParameters : propertyValue;
			} catch (Exception) {
				return extraParameters;
			}
		}

		/// <summary>
		/// Callback function to verify the server certificate.
		/// </summary>
		/// <param name="sender">An object that contains state information for this verification.</param>
		/// <param name="certificate">Certificate, used to verify the authenticity of the remote side.</param>
		/// <param name="chain">CA chain associated with the remote certificate.</param>
		/// <param name="policyErrors">One or more errors associated with the remote certificate.</param>
		/// <returns>Result command execution.</returns>
		public bool ValidateRemoteCertificate(object sender, X509Certificate certificate, X509Chain chain,
				SslPolicyErrors policyErrors) {
			return true;
		}

		public ExchangeCredentials GetCredentials(UserConnection userConnection, string senderEmailAddress) {
			var mailboxService = ClassFactory.Get<IMailboxService>(new ConstructorArgument("uc", userConnection));
			var mailbox = mailboxService.GetMailboxBySenderEmailAddress(senderEmailAddress);
			var credentials = new ExchangeCredentials {
				UserName = mailbox.Login,
				UserPassword = mailbox.Password,
				IsAutodiscover = mailbox.GetServerUseAutoDiscover(),
				ServerAddress = mailbox.GetServerAddress(),
				UseOAuth = mailbox.GetUseOAuth()
			};
			if (credentials.UseOAuth) {
				credentials.Token = mailbox.GetToken(userConnection);
			}
			return credentials;
		}

		#endregion

	}

	#endregion

	#region Struct: ExchangeMailFolder

	/// <summary>
	/// Ex#hange mail folder information wrapper.
	/// </summary>
	public struct ExchangeMailFolder
	{

		#region Properties: Public

		/// <summary>
		/// Unique identifier.
		/// </summary>
		public string Id { get; set; }

		/// <summary>
		/// Folder name.
		/// </summary>
		public string Name { get; set; }

		/// <summary>
		/// Parent folder unique identifier.
		/// </summary>
		public string ParentId { get; set; }

		/// <summary>
		/// Folder path string.
		/// </summary>
		public string Path { get; set; }

		/// <summary>
		/// If folder selected flag.
		/// </summary>
		public bool Selected { get; set; }

		#endregion

	}

	#endregion


	#region Class: ContactEmailAddressPropertiesMap

	/// <summary>
	/// ###### ###### ######## <see cref="Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition"/> ### #######
	/// ######## ## <see cref="Exchange.EmailAddressKey"/>. ######### ### ######## ######## #######
	/// <see cref="Exchange.Contact"/>.
	/// </summary>
	public class ContactEmailAddressPropertiesMap
	{

		#region Fields: Private

		private readonly Exchange.ExtendedPropertyDefinition[] _email1ExtendedPropertiesGroup =
		{
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email1AddressType, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email1DisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email1EmailAddress, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email1OriginalDisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email1OriginalEntryId, Exchange.MapiPropertyType.Binary)
		};

		private readonly Exchange.ExtendedPropertyDefinition[] _email2ExtendedPropertiesGroup =
		{
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email2AddressType, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email2DisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email2EmailAddress, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email2OriginalDisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email2OriginalEntryId, Exchange.MapiPropertyType.Binary)
		};

		private readonly Exchange.ExtendedPropertyDefinition[] _email3ExtendedPropertiesGroup =
		{
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email3AddressType, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email3DisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email3EmailAddress, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email3OriginalDisplayName, Exchange.MapiPropertyType.String),
			new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.Address,
				(int)EmailAddressExtendedProperties.Email3OriginalEntryId, Exchange.MapiPropertyType.Binary)
		};

		#endregion

		#region Methods: Public

		/// <summary>
		/// ########## ###### <see cref="Exchange.ExtendedPropertyDefinition"/> # ########### ## ########### <paramref name="emailKey"/>.
		/// </summary>
		/// <param name="emailKey">######## ############
		/// <see cref="Microsoft.Exchange.WebServices.Data.EmailAddressKey"/>.</param>
		/// <returns>###### <see cref="Exchange.ExtendedPropertyDefinition"/>.</returns>
		public Exchange.ExtendedPropertyDefinition[] GetExtendedPropertiesByKey(Exchange.EmailAddressKey emailKey) {
			Exchange.ExtendedPropertyDefinition[] result = null;
			switch (emailKey) {
				case Exchange.EmailAddressKey.EmailAddress1:
					result = _email1ExtendedPropertiesGroup;
					break;
				case Exchange.EmailAddressKey.EmailAddress2:
					result = _email2ExtendedPropertiesGroup;
					break;
				case Exchange.EmailAddressKey.EmailAddress3:
					result = _email3ExtendedPropertiesGroup;
					break;
			}
			return result;
		}

		#endregion

	}

	#endregion

	#region Enum: EmailAddressesExtendedProperties

	/// <summary>
	/// ########### Id ######## <see cref="Exchange.ExtendedPropertyDefinition"/>.
	/// ###### ######### ########## ## <seealso cref="http://msdn.microsoft.com/en-us/library/cc433490(v=exchg.80).aspx"/>.
	/// </summary>
	public enum EmailAddressExtendedProperties
	{
		Email1AddressType = 0x00008082,
		Email1DisplayName = 0x00008080,
		Email1EmailAddress = 0x00008083,
		Email1OriginalDisplayName = 0x00008084,
		Email1OriginalEntryId = 0x00008085,
		Email2AddressType = 0x00008092,
		Email2DisplayName = 0x00008090,
		Email2EmailAddress = 0x00008093,
		Email2OriginalDisplayName = 0x00008094,
		Email2OriginalEntryId = 0x00008095,
		Email3AddressType = 0x000080A2,
		Email3DisplayName = 0x000080A0,
		Email3EmailAddress = 0x000080A3,
		Email3OriginalDisplayName = 0x000080A4,
		Email3OriginalEntryId = 0x000080A5
	}

	#endregion

	#region Class: ExchangeDetailSynchronizer

	/// <summary>
	/// ######## ###### ############# ###### ####### <see cref="Exchange.Contact" /> # ########## <see cref="Contact" />.
	/// </summary>
	/// <typeparam name="TKey">#### ### ####### # ########### ####### # ####### ####### <see cref="Exchange.Contact" />.</typeparam>
	/// <typeparam name="TEntry">###### # ####### ####### <see cref="Exchange.Contact" />.</typeparam>
	/// <typeparam name="TRemoteItemType">### ######## Exchange.</typeparam>
	public abstract class ExchangeDetailSynchronizer<TKey, TEntry, TRemoteItemType>
		where TEntry : Exchange.DictionaryEntryProperty<TKey>
		where TRemoteItemType : Exchange.Item
	{

		#region Fields: Private

		private readonly SyncContext _context;
		private readonly LocalItem _localItem;
		private readonly string _detailItemTypeColumnName;
		private readonly string _detailSchemaName;

		#endregion

		#region Properties: Protected

		protected Exchange.DictionaryProperty<TKey, TEntry> DetailItems {
			get;
			set;
		}

		protected TRemoteItemType RemoteItem {
			get;
			private set;
		}

		protected SyncContext Context {
			get {
				return _context;
			}
		}

		protected LocalItem LocalItem {
			get {
				return _localItem;
			}
		}

		protected string DetailItemTypeColumnName {
			get {
				return _detailItemTypeColumnName;
			}
		}

		protected ExchangeUtilityImpl ExchangeUtility { get; } = new ExchangeUtilityImpl();

		#endregion

		#region Properties: Public

		public Dictionary<TKey, Guid> TypesMap {
			get;
			set;
		}

		#endregion

		#region Constructors: Protected

		protected ExchangeDetailSynchronizer(SyncContext context, LocalItem localItem,
			string detailItemTypeColumnName, TRemoteItemType remoteItem, string detailSchemaName) {
			_context = context;
			_localItem = localItem;
			_detailSchemaName = detailSchemaName;
			_detailItemTypeColumnName = detailItemTypeColumnName;
			RemoteItem = remoteItem;
		}

		#endregion

		#region Methods: Private

		private void AddNewDetailItemToLocalItem(TKey typeKey) {
			EntitySchema schema = Context.UserConnection.EntitySchemaManager.GetInstanceByName(_detailSchemaName);
			Entity detailEntity = schema.CreateEntity(Context.UserConnection);
			detailEntity.SetDefColumnValues();
			SetLocalItemValue(detailEntity, typeKey);
			DetailEntityConfig detailConfig = LocalItem.Schema.DetailConfigs.FirstOrDefault(
				config => config.SchemaName == _detailSchemaName);
			detailEntity.SetColumnValue(detailConfig.ForeingKeyColumnName + "Id",
				LocalItem.Entities[detailConfig.PrimarySchemaName][0].EntityId);
			SyncEntity newSyncEntity = SyncEntity.CreateNew(detailEntity);
			newSyncEntity.ExtraParameters = typeKey.ToString();
			LocalItem.AddOrReplace(_detailSchemaName, newSyncEntity);
		}

		private TKey ParseEnum(string value) {
			return (TKey)Enum.Parse(typeof(TKey), value, true);
		}

		private void InitializeDetailsCollections(out IEnumerable<SyncEntity> deletedItems,
			out IEnumerable<SyncEntity> existingItems) {
			List<SyncEntity> syncedEntities = (from detailEntity in LocalItem.Entities[_detailSchemaName]
											   where detailEntity.State != SyncState.Deleted
											   && TypesMap.Any(di => di.Key.ToString() == detailEntity.ExtraParameters)
											   select detailEntity).ToList();
			deletedItems = (from detailEntity in syncedEntities
							where !ContainsValue(ParseEnum(detailEntity.ExtraParameters))
							select detailEntity).ToList();
			IEnumerable<SyncEntity> deletedDetails = deletedItems;
			existingItems = (from detailEntity in syncedEntities
							 where deletedDetails.All(deleted =>
									 deleted.EntityId != detailEntity.EntityId)
							 select detailEntity).ToList();
		}

		private SyncEntity FindSyncedDetail(IEnumerable<SyncEntity> existingDetails, TKey typeKey) {
			return (from ed in existingDetails
					where ed.ExtraParameters == typeKey.ToString()
					select ed).FirstOrDefault();
		}

		#endregion

		#region Methods: Protected

		protected abstract bool ContainsValue(TKey typeKey);

		protected abstract void SetLocalItemValue(Entity detailItem, TKey typeKey);

		protected abstract void SetRemoteItemValue(Entity detailItem, TKey typeKey);

		protected abstract void DeleteRemoteDetail(TKey typeKey);

		#endregion

		#region Methods: Public

		/// <summary>
		/// ######### ######## ####### <see cref="Contact"/> ########## ####### <see cref="Exchange.Contact"/>.
		/// </summary>
		public void SyncLocalDetails() {
			IEnumerable<SyncEntity> deletedItems;
			IEnumerable<SyncEntity> existingItems;
			InitializeDetailsCollections(out deletedItems, out existingItems);
			foreach (KeyValuePair<TKey, Guid> typeMap in TypesMap) {
				TKey typeKey = typeMap.Key;
				if (ContainsValue(typeKey)) {
					SyncEntity syncedItem = (from item in existingItems
											 where item.ExtraParameters == typeKey.ToString()
											 select item).FirstOrDefault();
					//#### ####### ##### ################## ########
					if (syncedItem != null) {
						SetLocalItemValue(syncedItem.Entity, typeKey);
						syncedItem.Action = SyncAction.Update;
					} else {
						AddNewDetailItemToLocalItem(typeKey);
					}
				}
				//#############, #### ## ######### ######## # ##### typeKey
				SyncEntity deletedItem = (from item in deletedItems
										  where item.ExtraParameters == typeKey.ToString()
										  select item).FirstOrDefault();
				//#### ####, ####### ## ## remoteItem
				if (deletedItem != null) {
					deletedItem.Action = SyncAction.Delete;
				}
			}
		}

		/// <summary>
		/// ######### ######## ####### <see cref="Exchange.Contact"/> ########## ####### <see cref="Contact"/>.
		/// </summary>
		public void SyncRemoteDetails() {
			//######## ###### ############ #######
			IEnumerable<SyncEntity> existingSyncedDetails =
				from detailEntity in LocalItem.Entities[_detailSchemaName]
				where TypesMap.Any(dm => detailEntity.State !=
					SyncState.Deleted && dm.Key.ToString() == detailEntity.ExtraParameters)
				select detailEntity;
			//######## ###### ######### ######## #######
			IEnumerable<SyncEntity> deletedSyncedDetails =
				from detailEntity in LocalItem.Entities[_detailSchemaName]
				where TypesMap.Any(dm => detailEntity.State ==
						SyncState.Deleted && dm.Key.ToString() == detailEntity.ExtraParameters)
				select detailEntity;
			//######## ###### #################### #######
			List<SyncEntity> unsyncedItems = (from detailEntity in LocalItem.Entities[_detailSchemaName]
											  where string.IsNullOrEmpty(detailEntity.ExtraParameters)
											  orderby detailEntity.Entity.GetTypedColumnValue<DateTime>("CreatedOn")
											  select detailEntity).ToList();
			foreach (KeyValuePair<TKey, Guid> typeMap in TypesMap) {
				TKey typeKey = typeMap.Key;
				SyncEntity syncedItem = FindSyncedDetail(existingSyncedDetails, typeKey);
				//#### ####### ##### ################## ########
				if (syncedItem != null) {
					SetRemoteItemValue(syncedItem.Entity, typeKey);
				} else {
					SyncEntity unsyncedItem = unsyncedItems.FirstOrDefault(item =>
						item.Entity.GetTypedColumnValue<Guid>(_detailItemTypeColumnName) == typeMap.Value);
					if (unsyncedItem != null) {
						unsyncedItem.ExtraParameters = typeKey.ToString();
						SetRemoteItemValue(unsyncedItem.Entity, typeKey);
						unsyncedItems.Remove(unsyncedItem);
					}
				}
				//#############, #### ## ######### ######## # ##### typeKey
				SyncEntity deletedItem = FindSyncedDetail(deletedSyncedDetails, typeKey);
				if (deletedItem != null) {
					DeleteRemoteDetail(typeKey);
				}
			}
		}

		#endregion

	}

	#endregion

	public class TraceListenerInstance : Exchange.ITraceListener
	{

		private readonly ILog _log;
		private readonly string _mailbox;
		public TraceListenerInstance(ILog log, string mailbox) {
			_log = log;
			_mailbox = mailbox;
		}

		public void Trace(string traceType, string traceMessage) {
			_log.Debug($"[{_mailbox} - {traceType}] {traceMessage}");
		}
	}

}