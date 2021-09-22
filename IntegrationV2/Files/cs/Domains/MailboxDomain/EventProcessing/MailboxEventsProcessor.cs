namespace Terrasoft.MailboxDomain.EventProcessing
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
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;

	#region Class: MailboxEventsProcessor

	/// <summary>
	/// Mailbox events processor.
	/// </summary>
	[DefaultBinding(typeof(IMailboxEventsProcessor), Name = "MailboxEventsProcessor")]
	public class MailboxEventsProcessor : BaseEmailEventsProcessor, IMailboxEventsProcessor
	{

		#region Constructors: Public

		/// <summary>
		/// <see cref="MailboxEventsProcessor"/> ctor.
		/// </summary>
		/// <param name="appConnection"><see cref="AppConnection"/> instance.</param>
		/// <param name="accessor">see cref="IHttpContextAccessor"/> instance.</param>
		public MailboxEventsProcessor(AppConnection appConnection, IHttpContextAccessor accessor) : base(appConnection, accessor) {
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Updates existing exchange server email events subscription for <paramref name="mailboxId"/>.
		/// </summary>
		/// <param name="emailsData"><see cref="MailboxInfo"/> instance.</param>
		public void ProcessRefreshAccessToken(MailboxInfo mailboxInfo) {
			var userConnection = CreateUserConnection(mailboxInfo.BpmUserName, null, mailboxInfo.MailboxId);
			var managerFactory = ClassFactory.Get<IListenerManagerFactory>();
			var listenerManager = managerFactory.GetExchangeListenerManager(userConnection);
			listenerManager.UpdateListener(mailboxInfo.MailboxId, mailboxInfo.SenderEmailAddress);
		}

		#endregion

	}

	#endregion

}
