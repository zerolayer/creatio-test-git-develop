namespace Terrasoft.Mail.Sender
{
	using System;
	using System.Collections.Generic;
	using System.Net;
	using System.Net.Security;
	using System.Security.Cryptography.X509Certificates;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using global::Common.Logging;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using SysSettings = Terrasoft.Core.Configuration.SysSettings;
	using Terrasoft.Core.Factories;
	using MailServerModel = IntegrationApi.MailboxDomain.Model.MailServer;
	using IntegrationApi.MailboxDomain.Interfaces;

	#region Class: ExchangeClient

	/// <summary>Represents an Exchange mail client class.</summary>
	public class ExchangeClient : IEmailClient
	{

		#region Consts: Private

		/// <summary>
		/// PidTagInternetMessageId Canonical Property identifier.
		/// Corresponds to the message ID field as specified in [RFC2822].
		/// </summary>
		private const int _mapiMessageIdPropertyIdentifier = 4149;

		#endregion

		#region Constructors: Public

		/// <summary><see cref="ExchangeClient"/> ctor.</summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		public ExchangeClient(UserConnection userConnection) {
			_userConnection = userConnection;
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			_ignoreSslWarnings = (bool)SysSettings.GetValue(userConnection, "IgnoreExchangeSSLWarnings", false);
		}

		/// <summary><see cref="ExchangeClient"/> ctor.</summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		/// <param name="credentials">An instance of the user credentials.</param>
		public ExchangeClient(UserConnection userConnection, Credentials credentials)
			: this(userConnection) {
			var mailServer = GetMailServer(credentials.ServerId);
			_credentials = new ExchangeCredentials { 
				UserName = credentials.UserName,
				UserPassword = credentials.UserPassword,
				IsAutodiscover = mailServer.UseAutodiscover,
				ServerAddress = mailServer.ServerAddress,
				UseOAuth = credentials.UseOAuth
			};
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			_ignoreSslWarnings = (bool)SysSettings.GetValue(userConnection, "IgnoreExchangeSSLWarnings", false);
		}

		/// <summary><see cref="ExchangeClient"/> ctor.</summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		/// <param name="emailCredentials">An instance of the user email credentials.</param>
		public ExchangeClient(UserConnection userConnection, EmailContract.DTO.Credentials emailCredentials)
			: this(userConnection) {
			_credentials = new ExchangeCredentials { 
				UserName = emailCredentials.UserName,
				UserPassword = emailCredentials.Password,
				IsAutodiscover = emailCredentials.IsAutodiscover,
				ServerAddress = emailCredentials.ServiceUrl,
				UseOAuth = emailCredentials.UseOAuth
			};
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			_ignoreSslWarnings = (bool)SysSettings.GetValue(userConnection, "IgnoreExchangeSSLWarnings", false);
		}

		#endregion

		#region Fields: Private

		private readonly UserConnection _userConnection;
		private readonly ExchangeCredentials _credentials;
		private static ILog _log;
		private static bool _ignoreSslWarnings;
		private ExchangeUtilityImpl _exchangeUtility = new ExchangeUtilityImpl();

		#endregion

		#region Properties: Protected

		protected Exchange.ExchangeService Service { get; set; }

		#endregion

		#region Methods: Private

		private static bool ValidateRemoteCertificate(object sender, X509Certificate certificate, X509Chain chain,
				SslPolicyErrors policyErrors) {
			if (policyErrors == SslPolicyErrors.None || _ignoreSslWarnings) {
				return true;
			}
			if ((policyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0) {
				if (chain != null && chain.ChainStatus != null) {
					foreach (X509ChainStatus status in chain.ChainStatus) {
						if ((certificate.Subject == certificate.Issuer) && 
								(status.Status == X509ChainStatusFlags.UntrustedRoot)) {
							continue;
						} else if (status.Status != X509ChainStatusFlags.NoError) {
							string logTemplate = "ExchangeClient.ValidateRemoteCertificate -> "+
							" error RemoteCertificateChainErrors -> recived '{0}' chain status.\r\n" +
								"Status info: {1}\r\n Certificate: {2}, Subject {3}";
							Info(string.Format(logTemplate, status.Status, status.StatusInformation, 
								certificate.GetSerialNumberString(), certificate.Subject));
							return false;
						}
					}
				}
				return true;
			} else {
				return false;
			}
		}

		private IResourceStorage GetResourceStorage() {
			return _userConnection.Workspace.ResourceStorage;
		}

		private static Exchange.Importance GetExchangeImportance(EmailPriority emailPriority) {
			switch (emailPriority) {
				case EmailPriority.None:
				case EmailPriority.Normal:
					return Exchange.Importance.Normal;
				case EmailPriority.High:
					return Exchange.Importance.High;
				case EmailPriority.Low:
					return Exchange.Importance.Low;
				default:
					return Exchange.Importance.Normal;
			}
		}

		private Exchange.EmailMessage SetAttachments(Exchange.EmailMessage emailMessage, 
				IEnumerable<EmailAttachment> attachments) {
			foreach (EmailAttachment attachment in attachments) {
				var newAttachment = emailMessage.Attachments.AddFileAttachment(attachment.Name, attachment.Data);
				if (attachment.IsContent) {
					newAttachment.ContentId = attachment.Id.ToString();
					if (emailMessage.Body.Text.Contains("cid:" + attachment.Id)) {
						ActivityUtils.SetInlineFlagAtActivityFile(_userConnection, attachment.Id);
					}
				}
			}
			return emailMessage;
		}

		/// <summary>
		/// Create new instance of <see cref="Microsoft.Exchange.WebServices.Data.EmailMessage"/>
		/// based on data of <see cref="Terrasoft.Mail.Sender.EmailMessage"/> object.
		/// </summary>
		/// <param name="bpmEmailMessage">Email message data.</param>
		/// <param name="ignoreRights">Ignore rights when data requested flag.</param>
		/// <returns>Ready to send object of <see cref="Microsoft.Exchange.WebServices.Data.EmailMessage"/></returns>
		private Exchange.EmailMessage CreateExchangeEmailMessage(EmailMessage bpmEmailMessage, bool ignoreRights) {
			SetServiceConnection(bpmEmailMessage, ignoreRights);
			var exchangeEmailMessage = GetMessageInstance(bpmEmailMessage);
			exchangeEmailMessage = SetHeaderProperties(exchangeEmailMessage, bpmEmailMessage);
			exchangeEmailMessage = SetRecipients(exchangeEmailMessage, bpmEmailMessage);
			Info(string.Format("[ExchangeClient: {0} - {1}] Message \"{2}\" created", _userConnection.CurrentUser.Name,
				bpmEmailMessage.From, bpmEmailMessage.Subject));
			exchangeEmailMessage = SetAttachments(exchangeEmailMessage, bpmEmailMessage.Attachments);
			Info(string.Format("[ExchangeClient: {0} - {1}] Message \"{2}\" attachments added", _userConnection.CurrentUser.Name,
				bpmEmailMessage.From, bpmEmailMessage.Subject));
			return exchangeEmailMessage;
		}

		/// <summary>
		/// Sets <see cref="ExchangeClient.Service"/> property value.
		/// </summary>
		/// <param name="bpmEmailMessage">Email message instance.</param>
		/// <param name="ignoreRights">Ignore rights when data requested flag.</param>
		private void SetServiceConnection(EmailMessage bpmEmailMessage, bool ignoreRights) {
			string emailAddress = bpmEmailMessage.From.ExtractEmailAddress();
			if (Service == null) {
				if (_credentials != null) {
					Service = _exchangeUtility.CreateExchangeService(_userConnection, _credentials, emailAddress);
				} else {
					Service = _exchangeUtility.CreateExchangeService(_userConnection, emailAddress, false, ignoreRights);
				}
			}
		}

		/// <summary>
		/// Creates <see cref="Exchange.EmailMessage"/> instance using data from <paramref name="bpmEmailMessage"/>.
		/// </summary>
		/// <param name="bpmEmailMessage">Email message instance.</param>
		/// <returns><see cref="Exchange.EmailMessage"/> instance.</returns>
		private Exchange.EmailMessage GetMessageInstance(EmailMessage bpmEmailMessage) {
			string emailAddress = bpmEmailMessage.From.ExtractEmailAddress();
			var importance = GetExchangeImportance(bpmEmailMessage.Priority);
			var exchangeEmailMessage = new Exchange.EmailMessage(Service) {
					Subject = bpmEmailMessage.Subject,
					Body = StringUtilities.ReplaceInvalidXmlChars(bpmEmailMessage.Body),
					Importance = importance,
					From = emailAddress,
				};
			exchangeEmailMessage.Body.BodyType = bpmEmailMessage.IsHtmlBody ? Exchange.BodyType.HTML : Exchange.BodyType.Text;
			return exchangeEmailMessage;
		}

		/// <summary>
		/// Fills <paramref name="exchangeEmailMessage"/> headers collection using data from <paramref name="bpmEmailMessage"/>.
		/// </summary>
		/// <param name="exchangeEmailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="bpmEmailMessage">Email message instance.</param>
		/// <returns><see cref="Exchange.EmailMessage"/> instance.</returns>
		private Exchange.EmailMessage SetHeaderProperties(Exchange.EmailMessage exchangeEmailMessage, EmailMessage bpmMessage) {
			List<EmailMessageHeader> headerProperties = bpmMessage.HeaderProperties;
			if (headerProperties != null) {
				foreach (var property in headerProperties) {
					var header = new Exchange.ExtendedPropertyDefinition(Exchange.DefaultExtendedPropertySet.InternetHeaders,
						property.Name, Exchange.MapiPropertyType.String);
					exchangeEmailMessage.SetExtendedProperty(header, property.Value);
				}
			}
			var PidTagInternetMessageId = new Exchange.ExtendedPropertyDefinition(_mapiMessageIdPropertyIdentifier,
					Exchange.MapiPropertyType.String);
			exchangeEmailMessage.SetExtendedProperty(PidTagInternetMessageId, bpmMessage.GetMessageId());
			return exchangeEmailMessage;
		}

		/// <summary>
		/// Fills <paramref name="exchangeEmailMessage"/> recipients properties.
		/// </summary>
		/// <param name="exchangeEmailMessage"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="bpmEmailMessage">Email message instance.</param>
		/// <returns><see cref="Exchange.EmailMessage"/> instance.</returns>
		private Exchange.EmailMessage SetRecipients(Exchange.EmailMessage exchangeEmailMessage, EmailMessage bpmEmailMessage) {
			foreach (var recipient in bpmEmailMessage.To) {
				exchangeEmailMessage.ToRecipients.Add(recipient.ExtractEmailAddress());
			}
			foreach (var recipient in bpmEmailMessage.Cc) {
				exchangeEmailMessage.CcRecipients.Add(recipient.ExtractEmailAddress());
			}
			foreach (var recipient in bpmEmailMessage.Bcc) {
				exchangeEmailMessage.BccRecipients.Add(recipient.ExtractEmailAddress());
			}
			Info(string.Format("[ExchangeClient: {0} - {1}] Message \"{2}\" recipients added", _userConnection.CurrentUser.Name,
				bpmEmailMessage.From, bpmEmailMessage.Subject));
			return exchangeEmailMessage;
		}

		/// <summary>
		/// Returns mail server instance.</summary>
		/// <param name="serverId">Mail server identifier.</param>
		/// <returns>Mail server instance.</returns>
		private MailServerModel GetMailServer(Guid serverId) {
			var service = ClassFactory.Get<IMailServerService>(new ConstructorArgument("uc", _userConnection));
			return service.GetServer(serverId, false);
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Sends email message.
		/// </summary>
		/// <param name="exchangeEmailMessage">Prepared email message.</param>
		protected virtual void SendMessage(Exchange.EmailMessage exchangeEmailMessage) {
			try {
				string emailAddress = exchangeEmailMessage.From.Address.ExtractEmailAddress();
				var id = new Exchange.FolderId(Exchange.WellKnownFolderName.SentItems, emailAddress);
				exchangeEmailMessage.SendAndSaveCopy(id);
			} catch (Exception e) {
				Info(string.Format("[ExchangeClient: {0} - {1}] Error on send message \"{2}\", {3}", _userConnection.CurrentUser.Name,
					exchangeEmailMessage.From, exchangeEmailMessage.Subject, e.ToString()));
				throw;
			}
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Preparing email by <paramref name="emailMessage"/> data and send it.</summary>
		/// <param name="emailMessage">Email message data.</param>
		/// <param name="ignoreRights">Flag that indicates whether to ignore rights.</param>
		public void Send(EmailMessage emailMessage, bool ignoreRights) {
			Info(string.Format("[ExchangeClient: {0} - {1}] Start sending message \"{2}\"", _userConnection.CurrentUser.Name,
				emailMessage.From, emailMessage.Subject));
			Exchange.EmailMessage exchangeEmailMessage = CreateExchangeEmailMessage(emailMessage, ignoreRights);
			Info(string.Format("[ExchangeClient: {0} - {1}] Exchange item for message \"{2}\" created", _userConnection.CurrentUser.Name,
				emailMessage.From, emailMessage.Subject));
			exchangeEmailMessage.SetExtendedProperty(_exchangeUtility.GetContactExtendedPropertyDefinition(),
				_userConnection.CurrentUser.ContactId.ToString());
			Info(string.Format("[ExchangeClient: {0} - {1}] Exchange ExtendedProperty for message \"{2}\" created", _userConnection.CurrentUser.Name,
				emailMessage.From, emailMessage.Subject));
			SendMessage(exchangeEmailMessage);
			Info(string.Format("[ExchangeClient: {0} - {1}] Exchange item for message \"{2}\" sended", _userConnection.CurrentUser.Name,
				emailMessage.From, emailMessage.Subject));
		}
		
		public static void Info(string message) {
			if (_log == null) {
				_log = LogManager.GetLogger("Exchange");
			}
			_log.Info(message);
		}

		#endregion

	}

	#endregion
}