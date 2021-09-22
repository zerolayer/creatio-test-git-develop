namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Net;
	using System.Text;
	using EmailContract.DTO;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Common;
	using Terrasoft.Common.Json;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Logging.Interfaces;
	using Terrasoft.IntegrationV2.Utils;
	using MailboxFolder = EmailContract.DTO.MailboxFolder;
	using TS = Terrasoft.Web.Http.Abstractions;

	#region Class: ExchangeListenerManager

	/// <summary>
	/// Class provides methods for exchange listeners service interaction.
	/// </summary>
	[DefaultBinding(typeof(IExchangeListenerManager))]
	public class ExchangeListenerManager : IExchangeListenerManager
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="ISynchronizationLogger"/> instance.
		/// </summary>
		private readonly ISynchronizationLogger _log;

		#endregion

		#region Fields: Protected

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		protected readonly UserConnection UserConnection;

		/// <summary>
		/// <see cref="IHttpWebRequestFactory"/> implementation instance.
		/// </summary>
		protected readonly IHttpWebRequestFactory _requestFactory;

		/// <summary>
		/// <see cref="IMailboxService"/> implementation instance.
		/// </summary>
		protected readonly IMailboxService _mailboxService;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Initializes a new instance of the <see cref="ExchangeListenerManager"/> class.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		public ExchangeListenerManager(UserConnection userConnection) {
			UserConnection = userConnection;
			_requestFactory = ClassFactory.Get<IHttpWebRequestFactory>();
			_mailboxService = ClassFactory.Get<IMailboxService>(new ConstructorArgument("uc", userConnection));
			_log = ClassFactory.Get<ISynchronizationLogger>(new ConstructorArgument("userId", userConnection.CurrentUser.Id));
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// Current <see cref="TS.IHttpContextAccessor"/> instance.
		/// </summary>
		public TS.IHttpContextAccessor Context { get; set; } = TS.HttpContext.HttpContextAccessor != null ? TS.HttpContext.HttpContextAccessor : null;

		#endregion

		#region Method: Private

		/// <summary>
		/// Calls <paramref name="action"/> and handles thrown exceptions.
		/// </summary>
		/// <param name="action">Authentication action.</param>
		/// <param name="senderEmailAddress">Mailbox address.</param>
		private void TryDoListenerAction(string senderEmailAddress, Action action) {
			try {
				ListenerUtils.TryDoListenerAction(action, senderEmailAddress, UserConnection);
			} catch (Exception e) {
				var exceptionMessage = e.GetType() == typeof(AggregateException)
					? e.InnerException.Message
					: e.Message;
				_log.Error(exceptionMessage, e);
				throw;
			}
		}

		/// <summary>
		/// Returns is feature enabled for <paramref name="uc"/>.
		/// </summary>
		/// <param name="uc"><see cref="UserConnection"/> instance.</param>
		/// <param name="code">Feature code.</param>
		/// <returns><c>True</c> if feature enabled, otherwise returns false.</returns>
		private bool GetIsFeatureEnabled(string code) {
			return ListenerUtils.GetIsFeatureEnabled(UserConnection, code);
		}

		/// <summary>
		/// Gets exchange server email events subscription credentials.
		/// </summary>
		/// <param name="mailboxId"><see cref="MailboxSyncSettings"/> instance inique identifier.</param>
		/// <returns><see cref="SynchronizationCredentials"/> instance.</returns>
		private SynchronizationCredentials GetSynchronizationCredentials(string mailboxName) {
			var mailbox = _mailboxService.GetMailboxBySenderEmailAddress(mailboxName);
			if (mailbox == null) {
				throw new ItemNotFoundException($"Mailbox {mailboxName} not found");
			}
			return GetSynchronizationCredentials(mailbox);
		}

		/// <summary>
		/// Gets mail server <see cref="SynchronizationCredentials"/>.
		/// </summary>
		/// <param name="mailbox"><see cref="Mailbox"/> instance.</param>
		/// <returns><see cref="SynchronizationCredentials"/> instance.</returns>
		private SynchronizationCredentials GetSynchronizationCredentials(Mailbox mailbox) {
			var credentials = mailbox.ConvertToSynchronizationCredentials(UserConnection);
			var utils = ClassFactory.Get<ListenerUtils>(new ConstructorArgument("uc", UserConnection),
				new ConstructorArgument("context", Context));
			credentials.BpmEndpoint = utils.GetBpmEndpointUrl();
			return credentials;
		}

		/// <summary>
		/// Creates exchange server email events subscription parameters for <paramref name="mailboxId"/>.
		/// </summary>
		/// <param name="mailboxId"><see cref="MailboxSyncSettings"/> instance inique identifier.</param>
		/// <returns>Exchange server email events subscription parameters data array.</returns>
		private byte[] GetConnectionParams(Guid mailboxId) {
			var mailbox = _mailboxService.GetMailbox(mailboxId);
			if (mailbox == null) {
				throw new ItemNotFoundException($"Mailbox {mailboxId} not found");
			}
			var credentials = GetSynchronizationCredentials(mailbox);
			var json = Json.Serialize(credentials);
			return Encoding.UTF8.GetBytes($"[{json}]");
		}

		/// <summary>
		/// Serialize exchange server credentials.
		/// </summary>
		/// <param name="credentials"><see cref="SynchronizationCredentials"/> instance.</param>
		/// <returns>Exchange server credentials data array.</returns>
		private byte[] Serialize(SynchronizationCredentials credentials) {
			var json = Json.Serialize(credentials);
			return Encoding.UTF8.GetBytes(json);
		}

		/// <summary>
		/// Executes ecxchange listener action.
		/// </summary>
		/// <param name="data"><see cref="byte"/> array action parameters.</param>
		/// <param name="exchangeListenerAction"><see cref="ExchangeListenerActions"/> value.</param>
		/// <returns>Response string value.</returns>
		private string ExecuteListenerAction(byte[] data, string exchangeListenerAction) {
			string serviceUri = ExchangeListenerActions.GetActionUrl(UserConnection, exchangeListenerAction);
			return ExecuteAction(data, serviceUri);
		}

		/// <summary>
		/// Returns mailbox folders list from integration provider.
		/// </summary>
		/// <param name="data"><see cref="byte"/> array action parameters.</param>
		/// <returns>Response string value.</returns>
		private string GetFoldersFromProvider(byte[] data) {
			string serviceUri = ExchangeListenerActions.GetAllMailboxFoldersUrl(UserConnection);
			return ExecuteAction(data, serviceUri);
		}

		/// <summary>
		/// Send <paramref name="data"/> to <paramref name="serviceUri"/>.
		/// </summary>
		/// <param name="data">Sending data.</param>
		/// <param name="serviceUri">Endpoint to data send.</param>
		/// <returns>Response string value.</returns>
		private string ExecuteAction(byte[] data, string serviceUri) {
			WebRequest request = _requestFactory.Create(serviceUri);
			request.Method = "POST";
			request.ContentType = "application/json; charset=utf-8";
			request.ContentLength = data.Length;
			request.Timeout = 5 * 60 * 1000;
			using (Stream stream = request.GetRequestStream()) {
				stream.Write(data, 0, data.Length);
			}
			WebResponse response = request.GetResponse();
			try {
				using (Stream dataStream = response.GetResponseStream()) {
					StreamReader reader = new StreamReader(dataStream);
					return reader.ReadToEnd();
				}
			} finally { 
				response.Close();
			}
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IExchangeListenerManager.StartListener"/>
		public void StartListener(Guid mailboxId) {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return;
			}
			var mailbox = _mailboxService.GetMailbox(mailboxId);
			if (!mailbox.CheckSynchronizationSettings()) {
				_log.Warn($"mailbox {mailbox.SenderEmailAddress} synchronization settings not valid");
				return;
			}
			byte[] data = GetConnectionParams(mailboxId);
			TryDoListenerAction(mailbox.SenderEmailAddress, () => {
				ExecuteListenerAction(data, ExchangeListenerActions.Create);
			});
		}

		/// <inheritdoc cref="IExchangeListenerManager.StopListener"/>
		public void StopListener(Guid mailboxId) {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return;
			}
			byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(new[] { mailboxId }));
			var mailbox = _mailboxService.GetMailbox(mailboxId);
			TryDoListenerAction(mailbox.SenderEmailAddress, () => {
				ExecuteListenerAction(data, ExchangeListenerActions.Close);
			});
		}

		/// <inheritdoc cref="IExchangeListenerManager.RecreateListener"/>
		public void RecreateListener(Guid mailboxId) {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return;
			}
			var mailbox = _mailboxService.GetMailbox(mailboxId);
			if (!mailbox.CheckSynchronizationSettings()) {
				StopListener(mailboxId);
				_log.Warn($"mailbox {mailbox.SenderEmailAddress} synchronization settings not valid");
				return;
			}
			byte[] data = GetConnectionParams(mailboxId);
			TryDoListenerAction(mailbox.SenderEmailAddress, () => {
				ExecuteListenerAction(data, ExchangeListenerActions.Recreate);
			});
		}

		/// <inheritdoc cref="IExchangeListenerManager.UpdateListener"/>
		public void UpdateListener(Guid mailboxId, string senderEmailAddress) {
			byte[] data = GetConnectionParams(mailboxId);
			TryDoListenerAction(senderEmailAddress, () => {
				ExecuteListenerAction(data, ExchangeListenerActions.Update);
			});
		}

		/// <inheritdoc cref="IExchangeListenerManager.GetIsServiceAvaliable"/>
		public bool GetIsServiceAvaliable() {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return false;
			}
			HttpWebResponse response = null;
			try {
				string serviceUri = ExchangeListenerActions.GetActionUrl(UserConnection, ExchangeListenerActions.Exists);
				WebRequest request = _requestFactory.Create(serviceUri);
				request.ContentType = "application/json; charset=utf-8";
				request.Timeout = 5 * 60 * 1000;
				response = (HttpWebResponse)request.GetResponse();
				return response.StatusCode == HttpStatusCode.OK;
			} catch (Exception) {
				return false;
			} finally {
				response?.Close();
			}
		}

		/// <inheritdoc cref="IExchangeListenerManager.GetSubscriptionsStatuses"/>
		public Dictionary<Guid, string> GetSubscriptionsStatuses(Guid[] mailboxIds) {
			var result = new Dictionary<Guid, string>();
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return result;
			}
			byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(mailboxIds));
			var rawResult = ExecuteListenerAction(data, ExchangeListenerActions.SubscriptionsState);
			JObject resultObj = Json.Deserialize(rawResult) as JObject;
			foreach (var item in resultObj) {
				Guid key = Guid.Parse(item.Key);
				if (!result.ContainsKey(key)) {
					result.Add(key, item.Value.Value<string>("State"));
				}
			}
			return result;
		}

		/// <inheritdoc cref="IExchangeListenerManager.ValidateCredentials"/>
		public CredentialsValidationInfo ValidateCredentials(Mailbox mailbox) {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return new CredentialsValidationInfo {
					IsValid = false,
					Message = "Feature OldEmailIntegration enabled"
				};
			}
			var credentials = GetSynchronizationCredentials(mailbox);
			var data = Serialize(credentials);
			var rawResult = string.Empty;
			TryDoListenerAction(mailbox.SenderEmailAddress, () => {
				rawResult = ExecuteListenerAction(data, ExchangeListenerActions.Validate);
			});
			return Json.Deserialize<CredentialsValidationInfo>(rawResult);
		}

		/// <inheritdoc cref="IExchangeListenerManager.GetMailboxFolders"/>
		public IEnumerable<MailboxFolder> GetMailboxFolders(string mailboxName, string folderClassName = "") {
			if (GetIsFeatureEnabled("OldEmailIntegration")) {
				return new List<MailboxFolder>();
			}
			var credentials = GetSynchronizationCredentials(mailboxName);
			credentials.FolderClassName = folderClassName;
			var data = Serialize(credentials);
			var rawResult = string.Empty;
			TryDoListenerAction(mailboxName, () => {
				rawResult = GetFoldersFromProvider(data);
			});
			return Json.Deserialize<IEnumerable<MailboxFolder>>(rawResult);
		}

		#endregion

	}

	#endregion

}