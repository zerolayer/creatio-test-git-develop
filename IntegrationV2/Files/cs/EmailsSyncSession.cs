namespace Terrasoft.EmailDomain
{
	using System;
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Logging.Interfaces;
	using Terrasoft.IntegrationV2.Utils;

	#region Class: EmailsSyncSession

	[DefaultBinding(typeof(ISyncSession), Name = "Email")]
	public class EmailsSyncSession : ISyncSession
	{

		#region Fields: Private

		private readonly UserConnection _userConnection;

		private readonly string _senderEmailAddress;

		private readonly ISynchronizationLogger _log;

		#endregion

		#region Constructors: Public

		public EmailsSyncSession(UserConnection uc, string senderEmailAddress) {
			_log = ClassFactory.Get<ISynchronizationLogger>(new ConstructorArgument("userId", uc.CurrentUser.Id));
			_userConnection = uc;
			_senderEmailAddress = senderEmailAddress;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Returns current session mailbox model.
		/// </summary>
		/// <returns><see cref="Mailbox"/> instance.</returns>
		private Mailbox GetMailbox() {
			var mailboxService = ClassFactory.Get<IMailboxService>(new ConstructorArgument("uc", _userConnection));
			return mailboxService.GetMailboxBySenderEmailAddress(_senderEmailAddress);
		}

		/// <summary>
		/// Returns current session mailbox creadentials.
		/// </summary>
		/// <param name="mailbox"><see cref="Mailbox"/> instance.</param>
		/// <returns><see cref="SynchronizationCredentials"/> instance.</returns>
		private SynchronizationCredentials GetCredentials(Mailbox mailbox) {
			var credentials = mailbox.ConvertToSynchronizationCredentials(_userConnection);
			var utils = ClassFactory.Get<ListenerUtils>(new ConstructorArgument("uc", _userConnection),
				new ConstructorArgument("context", null));
			credentials.BpmEndpoint = utils.GetBpmEndpointUrl();
			return credentials;
		}

		/// <summary>
		/// Returns <see cref="IEmailProvider"/> implementation instance.
		/// </summary>
		/// <returns><see cref="IEmailProvider"/> implementation instance.</returns>
		private IEmailProvider GetProvider() {
			var requestFactory = ClassFactory.Get<IHttpWebRequestFactory>();
			return ClassFactory.Get<IEmailProvider>(
				new ConstructorArgument("userConnection", _userConnection),
				new ConstructorArgument("requestFactory", requestFactory));
		}

		/// <summary>
		/// Starts mailbox synchronization.
		/// </summary>
		/// <param name="mailbox"><see cref="Mailbox"/> instance.</param>
		/// <param name="filters">Synchronization session filters.</param>
		private void StartSynchronization(Mailbox mailbox, string filters) {
			if (!mailbox.CheckSynchronizationSettings()) {
				_log.Warn($"mailbox {mailbox.SenderEmailAddress} synchronization settings not valid");
				return;
			}
			var credentials = GetCredentials(mailbox);
			var emailProvider = GetProvider();
			Action action = () => {
				emailProvider.StartSynchronization(credentials, filters);
			};
			try {
				ListenerUtils.TryDoListenerAction(action, credentials.SenderEmailAddress, _userConnection);
			} catch(Exception e) {
				_log.Error($"Synchronization of {_senderEmailAddress} failed", e);
				throw;
			}
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="ISyncSession.Start"/>.
		public void Start() {
			_log.DebugFormat($"Synchronization of {_senderEmailAddress} started");
			var mailbox = GetMailbox();
			var filters = mailbox.GetFilters(_userConnection);
			StartSynchronization(mailbox, filters);
			_log.DebugFormat($"Synchronization of {_senderEmailAddress} initialization ended");
		}

		/// <inheritdoc cref="ISyncSession.StartFailover"/>.
		public void StartFailover() {
			_log.DebugFormat($"Failover synchronization of {_senderEmailAddress} started");
			var mailbox = GetMailbox();
			var sinceDate = ListenerUtils.GetFailoverPeriodStartDate(mailbox, _userConnection);
			var filters = mailbox.GetFailoverFilters(_userConnection, sinceDate);
			StartSynchronization(mailbox, filters);
			_log.DebugFormat($"Failover synchronization of {_senderEmailAddress} initialization ended");
		}

		#endregion

	}

	#endregion

}
