namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using IntegrationApi.Interfaces;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Logging.Interfaces;

	#region Class: ListenerServiceFailJob

	/// <summary>
	/// Minutely cheks the subscription statuses on the microservice, and creates immediate job <see cref="ListenerServiceFailHandler"/> 
	/// for mailboxes that does not have subscription.
	/// </summary>
	public class ListenerServiceFailJob: IJobExecutor
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="ISynchronizationLogger"/> instance.
		/// </summary>
		private ISynchronizationLogger _log;

		/// <summary>
		/// Mailbox failure handlers job group name.
		/// </summary>
		private readonly string _jobGroupName = "ExchangeListenerHandler";

		/// <summary>
		/// Mailbox subscription exists on litener service state code.
		/// </summary>
		private readonly string _subscriptionExistsState = "exists";

		/// <summary>
		/// <see cref="ExchangeListenerManager"/> instance.
		/// </summary>
		private IExchangeListenerManager ListenerManager;

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private UserConnection UserConnection;

		#endregion

		#region Methods: Private

		/// <summary>
		/// Checks existing mailboxes subscriptions state. Returns mailboxes without subscriptions.
		/// </summary>
		private List<Mailbox> GetMailboxesWithoutSubscriptions( ) {
			_log.DebugFormat("GetMailboxesWithoutSubscriptions method started");
			var mailboxes = GetSynchronizableMailboxes();
			_log.InfoFormat("GetMailboxesWithoutSubscriptions: selected {0} mailboxes", mailboxes.Count);
			if (!NeedProceed(mailboxes)) {
				_log.DebugFormat("GetMailboxesWithoutSubscriptions method ended");
				return new List<Mailbox>();
			}
			if (GetIsListenerServiceAvaliable()) {
				_log.DebugFormat("GetMailboxesWithoutSubscriptions: listener service avaliable, mailboxes subscriptions check started");
				mailboxes = FilterActiveMailboxes(mailboxes);
				_log.InfoFormat("GetMailboxesWithoutSubscriptions: filtered to {0} mailboxes", mailboxes.Count);
			}
			if (!NeedProceed(mailboxes)) {
				_log.DebugFormat("GetMailboxesWithoutSubscriptions method ended");
				return new List<Mailbox>();
			}
			_log.DebugFormat("GetMailboxesWithoutSubscriptions method ended");
			return mailboxes;
		}

		/// <summary>
		/// Starts failover handlers for <paramref name="mailboxes"/>.
		/// </summary>
		/// <param name="mailboxes">Mailboxes without subscriptions collection.</param>
		private void ScheduleFailoverHandlers(List<Mailbox> mailboxes) {
			_log.DebugFormat("StartFailoverHandlers started");
			var schedulerWraper = ClassFactory.Get<IAppSchedulerWraper>();
			schedulerWraper.RemoveGroupJobs(_jobGroupName);
			_log.DebugFormat("All jobs from {0} job group deleted.", _jobGroupName);
			foreach (var mailbox in mailboxes) {
				var parameters = new Dictionary<string, object> {
					{ "SenderEmailAddress", mailbox.SenderEmailAddress },
					{ "MailboxType", mailbox.TypeId },
					{ "MailboxId", mailbox.Id },
					{ "PeriodInMinutes", 0 }
				};
				var syncJobScheduler = ClassFactory.Get<ISyncJobScheduler>();
				if (!syncJobScheduler.DoesSyncJobExist(UserConnection, parameters)) {
					schedulerWraper.ScheduleImmediateJob<ListenerServiceFailHandler>(_jobGroupName, UserConnection.Workspace.Name,
						mailbox.OwnerUserName, parameters);
					_log.DebugFormat("ListenerServiceFailHandler for {0} mailbox started using {1} user.", mailbox.Id, mailbox.OwnerUserName);
				} else {
					_log.DebugFormat("ListenerServiceFailHandler for {0} mailbox skipped using {1} user.", mailbox.Id, mailbox.OwnerUserName);
				}
			}
			_log.DebugFormat("StartFailoverHandlers ended");
		}

		/// <summary>
		/// Checks is feature enabled for any user.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="code">Feature code.</param>
		/// <returns><c>True</c> if feature enabled for any user. Returns <c>false</c> otherwise.</returns>
		private bool GetIsFeatureEnabledForAnyUser(UserConnection userConnection, string code) {
			var featureUtil = ClassFactory.Get<IFeatureUtilities>();
			return featureUtil.GetIsFeatureEnabledForAnyUser(userConnection, code);
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns mailboxes with enabled email synchronization.
		/// </summary>
		/// <returns>Collection of mailboxes with enabled email synchronization.</returns>
		protected List<Mailbox> GetSynchronizableMailboxes() {
			_log.DebugFormat("GetAllSynchronizableMailboxes started");
			var mailboxService = ClassFactory.Get<IMailboxService>(new ConstructorArgument("uc", UserConnection));
			var result = mailboxService.GetAllSynchronizableMailboxes();
			_log.DebugFormat("GetAllSynchronizableMailboxes ended");
			return result;
		}

		/// <summary>
		/// Checks that exchange listeners service avaliable.
		/// </summary>
		/// <returns><c>True</c> if exchange listeners service avaliable. Otherwise returns <c>false</c>.</returns>
		protected bool GetIsListenerServiceAvaliable() {
			return ListenerManager.GetIsServiceAvaliable();
		}

		/// <summary>
		/// Removes mailboxes with active subscriptions from <paramref name="mailboxes"/>.
		/// </summary>
		/// <param name="mailboxes">Collection of exchange mailboxes with enabled email synchronization.</param>
		protected List<Mailbox> FilterActiveMailboxes(List<Mailbox> mailboxes) {
			_log.DebugFormat("FilterActiveMailboxes started");
			var subscriptions = ListenerManager.GetSubscriptionsStatuses(mailboxes.Select(kvp => kvp.Id).ToArray());
			_log.DebugFormat($"FilterActiveMailboxes ended. Recived {subscriptions.Count} subscriptions from listener service");
			var existingSubscriptions = subscriptions.Where(kvp => kvp.Value == _subscriptionExistsState).Select(kvp => kvp.Key);
			return mailboxes.Where(m => !existingSubscriptions.Contains(m.Id)).ToList();
		}

		/// <summary>
		/// Creates <see cref="ExchangeListenerManager"/> instance.
		/// </summary>
		protected IExchangeListenerManager GetExchangeListenerManager() {
			var managerFactory = ClassFactory.Get<IListenerManagerFactory>();
			return managerFactory.GetExchangeListenerManager(UserConnection);
		}

		/// <summary>
		/// Returns is fail handling steel needed for <paramref name="mailboxes"/>.
		/// </summary>
		/// <param name="mailboxes">Collection of exchange mailboxes with enabled email synchronization.</param>
		/// <returns><c>True</c> if fail handling steel needed for <paramref name="mailboxes"/>. 
		/// Otherwise returns <c>false</c>.</returns>
		protected bool NeedProceed(List<Mailbox> mailboxes) {
			return mailboxes.Count > 0;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Check exchange mailboxes subscriptions state. Starts exchange events service. 
		/// Fail handler for mailboxes without subscription.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Process parameters collection.</param>
		public void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			_log = ClassFactory.Get<ISynchronizationLogger>(new ConstructorArgument("userId", userConnection.CurrentUser.Id));
			try {
				_log.InfoFormat("ListenerServiceFailJob started");
				if (GetIsFeatureEnabledForAnyUser(userConnection, "OldEmailIntegration")) {
					return;
				}
				UserConnection = userConnection;
				ListenerManager = GetExchangeListenerManager();
				_log.DebugFormat("ListenerServiceFailJob: _listenerManager initiated");
				var mailboxes = GetMailboxesWithoutSubscriptions();
				_log.InfoFormat("ListenerServiceFailJob: mailboxes recived");
				ScheduleFailoverHandlers(mailboxes);
				_log.InfoFormat("ListenerServiceFailJob: Failover handlers scheduled");
			} catch (Exception e) {
				_log.Error($"ListenerServiceFailJob error {e}", e);
			} finally {
				int periodMin = Core.Configuration.SysSettings.GetValue(userConnection, "ListenerServiceFailJobPeriod", 1);
				if (periodMin == 0) {
					var schedulerWraper = ClassFactory.Get<IAppSchedulerWraper>();
					schedulerWraper.RemoveGroupJobs(ListenerServiceFailJobFactory.JobGroupName);
					_log.ErrorFormat("ListenerServiceFailJobPeriod is 0, ListenerServiceFailJob stopped");
				}
				_log.InfoFormat("ListenerServiceFailJob ended");
			}
		}

		#endregion

	}

	#endregion

}