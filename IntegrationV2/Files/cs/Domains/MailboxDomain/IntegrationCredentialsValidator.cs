namespace IntegrationV2.MailboxDomain
{
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using IntegrationApi.MailboxDomain;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;

	#region Class: IntegrationCredentialsValidator

	[DefaultBinding(typeof(IMailboxValidator), Name = "IntegrationCredentialsValidator")]
	public class IntegrationCredentialsValidator : BaseMailboxValidator, IMailboxValidator
	{

		#region Fields: Private

		private IExchangeListenerManager _listenerManager;

		#endregion
		
		#region Constructors: Public

		public IntegrationCredentialsValidator(UserConnection uc): base(uc) {
			var listenerManagerFactory = ClassFactory.Get<IListenerManagerFactory>();
			_listenerManager = listenerManagerFactory.GetExchangeListenerManager(uc);
		}

		#endregion

		#region Methods: Public

		///<inheritdoc cref="IMailboxValidator.ValidateSynchronization(Mailbox)"/>
		public CredentialsValidationInfo ValidateSynchronization(Mailbox mailbox) {
			return _listenerManager.ValidateCredentials(mailbox);
		}

		///<inheritdoc cref="IMailboxValidator.ValidateEmailSend(Mailbox)"/>
		public CredentialsValidationInfo ValidateEmailSend(Mailbox mailbox) {
			return SendTestMessage(mailbox);
		}

		#endregion

	}

	#endregion

}
