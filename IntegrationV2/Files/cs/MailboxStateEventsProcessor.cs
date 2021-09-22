namespace IntegrationV2
{
	using System;
	using System.Collections.Generic;
	using global::Common.Logging;
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core;

	#region Class: EmailEventsProcessor

	/// <summary>
	/// Mailbox state event handler.
	/// </summary>
	[DefaultBinding(typeof(IMailboxStateEventsProcessor))]
	public class MailboxStateEventsProcessor : IMailboxStateEventsProcessor
	{

		#region Fields: Private

		private readonly ILog _log = LogManager.GetLogger("ExchangeListener");

		#endregion

		#region Methods: Public

		/// <summary>
		/// Processes mailbox status from exchange listener service.
		/// </summary>
		/// <param name="mailboxState"><see cref="MailboxState"/> instance.</param>
		public void ProcessMailboxState(MailboxState mailboxState) {
			_log.Info($"ProcessConnectionStatus for {mailboxState.SenderEmailAddress} started.");
			var parameters = new Dictionary<string, object> {
				{ "SenderEmailAddress", mailboxState.SenderEmailAddress },
				{ "Avaliable", mailboxState.Avaliable },
				{ "ExceptionClassName", mailboxState.ExceptionClassName },
				{ "ExceptionMessage", mailboxState.ExceptionMessage },
				{ "Context", mailboxState.Context }
			};
			IAppSchedulerWraper schedulerWraper = ClassFactory.Get<IAppSchedulerWraper>();
			schedulerWraper.ScheduleImmediateJob<MailboxStateEventExecutor>(
					$"ExchangeListerProcessConnectionStatus_{mailboxState.SenderEmailAddress}_{Guid.NewGuid().ToString()}",
					mailboxState.BpmUserWorkspace, mailboxState.BpmUserName, parameters, true);
			_log.Info($"ProcessConnectionStatus for {mailboxState.SenderEmailAddress} ended.");
		}

		#endregion

	}

	#endregion

}
