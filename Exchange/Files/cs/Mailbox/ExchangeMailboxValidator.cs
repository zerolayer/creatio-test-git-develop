namespace Terrasoft.Configuration
{
	using System;
	using System.Linq;
	using System.Net;
	using System.Net.Security;
	using System.Security.Cryptography.X509Certificates;
	using EmailContract.DTO;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using IntegrationApi.MailboxDomain;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Newtonsoft.Json;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.Mail.Sender;

	#region Class: ExchangeMailboxValidator

	/// <summary>
	/// EWS library exchange mailbox validator implementation.
	/// </summary>
	[DefaultBinding(typeof(IMailboxValidator), Name = "ExchangeMailboxValidator")]
	public class ExchangeMailboxValidator : BaseMailboxValidator, IMailboxValidator
	{

		#region Fields: Private

		private ExchangeUtilityImpl _exchangeUtility = new ExchangeUtilityImpl();

		private static bool _ignoreSslWarnings;

		#endregion

		#region Constructors: Public
		
		public ExchangeMailboxValidator(UserConnection uc): base(uc) {
		}

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
							return false;
						}
					}
				}
				return true;
			} else {
				return false;
			}
		}

		#endregion

		#region Methods: Protected

		/// <inheritdoc cref="BaseMailboxValidator.GetEmailClientFactory"/>
		protected override IEmailClientFactory GetEmailClientFactory() {
			return ClassFactory.Get<EmailClientFactory>(new ConstructorArgument("userConnection",
				UserConnection));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IMailboxValidator.ValidateSynchronization"/>
		public CredentialsValidationInfo ValidateSynchronization(Mailbox mailbox) {
			var answer = new CredentialsValidationInfo() {
				IsValid = true
			};
			var credentials = mailbox.ConvertToSynchronizationCredentials(UserConnection);
			ServicePointManager.ServerCertificateValidationCallback += ValidateRemoteCertificate;
			_ignoreSslWarnings = (bool)Terrasoft.Core.Configuration.SysSettings.GetValue(UserConnection, "IgnoreExchangeSSLWarnings");
			try {
				var service = _exchangeUtility.CreateExchangeService(UserConnection, credentials,
					credentials.SenderEmailAddress);
				Exchange.NameResolutionCollection resolutions = service.ResolveName(credentials.SenderEmailAddress);
				var mailboxName = string.Empty;
				if (resolutions.Any()) {
					Exchange.EmailAddress mailboxAddress = resolutions.First().Mailbox;
					mailboxName = mailboxAddress.Name;
				}
				answer.Data = JsonConvert.SerializeObject(new {
					MailboxName = mailboxName
				});
			} catch (Exception exception) {
				answer.IsValid = false;
				answer.Message = ConnectToServerCaption + exception.Message;
			} finally {
				ServicePointManager.ServerCertificateValidationCallback -= ValidateRemoteCertificate;
			}
			return answer;
		}

		/// <inheritdoc cref="IMailboxValidator.ValidateSynchronization"/>
		public CredentialsValidationInfo ValidateEmailSend(Mailbox mailbox) {
			return SendTestMessage(mailbox);
		}

		#endregion

	}

	#endregion

}
