namespace Terrasoft.EmailDomain.EventProcessing
{
	using System;
	using System.Collections.Generic;
	using EmailContract.Commands;
	using IntegrationApi.Email;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.Web.Http.Abstractions;
	using Terrasoft.IntegrationV2.Logging.Interfaces;
	using IntegrationApi.MailboxDomain;

	#region Class: EmailEventsProcessor

	/// <summary>
	/// Load emails command handler.
	/// </summary>
	[DefaultBinding(typeof(IEmailEventsProcessor))]
	public class EmailEventsProcessor : BaseEmailEventsProcessor, IEmailEventsProcessor
	{

		#region Fields: Private

		private ISynchronizationLogger _log;

		#endregion

		#region Constructors: Public 

		public EmailEventsProcessor(AppConnection appConnection, IHttpContextAccessor accessor) : base(appConnection, accessor) {
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Checks is <paramref name="mailboxId"/> avaliable for current user.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="mailboxId">Mailbox identifier.</param>
		private void CheckMailboxAvaliable(UserConnection userConnection, Guid mailboxId) {
			GetMailboxAddress(userConnection, mailboxId);
		}

		#endregion

		#region Methods: Protected


		/// <summary>
		/// Starts <paramref name="emailsData"/> synchronization.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="emailsData"><see cref="LoadEmailCommand"/> instance.</param>
		protected void StartSynchronization(UserConnection userConnection, LoadEmailCommand emailsData) {
			Guid mailboxSyncSettingsId = emailsData.SubscriptionInfo.MailboxId;
			CheckMailboxAvaliable(userConnection, mailboxSyncSettingsId);
			var parameters = new Dictionary<string, object> {
				{ "MailboxId", mailboxSyncSettingsId },
				{ "Items", emailsData.Emails }
			};
			_log.Info($"StartSynchronization mailbox {mailboxSyncSettingsId} avaliable.");
			ValidateEvent(userConnection, parameters);
			_log.Info($"StartSynchronization mailbox {mailboxSyncSettingsId} valid.");
			Terrasoft.Core.Tasks.Task.StartNewWithUserConnection<LoadEmailEventExecutor, IDictionary<string, object>>(parameters);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Creates user connection and starts <paramref name="emailsData"/> synchronization.
		/// </summary>
		/// <param name="emailsData"><see cref="LoadEmailCommand"/> instance.</param>
		public void ProcessLoadEmailCommand(LoadEmailCommand emailsData) {
			var userName = emailsData.SubscriptionInfo.SysAdminUnit;
			var subscriptionInfo = emailsData.SubscriptionInfo;
			var userConnection = CreateUserConnection(userName, subscriptionInfo.TimeZoneId, subscriptionInfo.MailboxId);
			_log = ClassFactory.Get<ISynchronizationLogger>(new ConstructorArgument("userId", userConnection.CurrentUser.Id));
			_log.Info($"UserConnection for {emailsData.SubscriptionInfo.SysAdminUnit} created.");
			try {
				StartSynchronization(userConnection, emailsData);
			} catch (MailboxNotFountException e) {
				_log.Error(e.Message, e);
			}
			finally {
				userConnection?.Close(SessionEndMethod.Logout, false);
			}
			_log.Info($"ProcessLoadEmailCommand for {emailsData.SubscriptionInfo.MailboxId} ended.");
		}

		#endregion

	}

	#endregion

}
