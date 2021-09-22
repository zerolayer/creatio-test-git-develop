namespace IntegrationV2
{
	using System;
	using System.Collections.Generic;
	using Common.Logging;
	using IntegrationApi.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;

	#region Class: SyncJobScheduler

	[DefaultBinding(typeof(ISyncJobScheduler))]
	public class SyncJobScheduler : ISyncJobScheduler
	{

		#region Fields: Private

		/// <summary>
		/// Synchronization jobs name.
		/// </summary>
		private const string SyncJobGroupName = "Exchange";
		
		/// <summary>
		/// Exchange integration logger.
		/// </summary>
		private ILog _log { get; } = LogManager.GetLogger("Exchange");

		/// <summary>
		/// <see cref="IAppSchedulerWraper"/> implementation.
		/// </summary>
		private readonly IAppSchedulerWraper _appSchedulerWraper = ClassFactory.Get<IAppSchedulerWraper>();

		/// <summary>
		/// Default synchronization job suffix.
		/// </summary>
		private readonly string _syncJobSuffix = "ImmediateProcessJob";

		/// <summary>
		/// Allowed syncronization job parameters for <see cref="IAppSchedulerWraper"/>.
		/// </summary>
		private readonly List<string> _syncJobParameters = new List<string>{
			"SenderEmailAddress",
			"LoadEmailsFromDate",
			"MailboxId"
		};

		#endregion

		#region Methods: Private

		/// <summary>
		/// Get simply synchronization job name.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="syncProcessName">Sync process name.</param>
		/// <returns>Synchronization job name.</returns>
		private string GetSyncJobName(UserConnection userConnection, string syncProcessName) {
			return $"{syncProcessName}_{userConnection.CurrentUser.Id}";
		}

		/// <summary>
		/// Get synchronization job name.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="processName">Process name.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="suffix">Process name suffix.</param>
		/// <returns>Synchronization job name.</returns>
		private string GetSyncJobName(UserConnection userConnection, string processName,
				string senderEmailAddress, string suffix = null) {
			string result = $"{senderEmailAddress}_{processName}_{userConnection.CurrentUser.Id}";
			return string.IsNullOrEmpty(suffix) ? result : $"{result}_{suffix}";
		}

		/// <summary>
		/// Get synchronization job name.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <returns>Synchronization job name.</returns>
		private string GetSyncJobName(UserConnection userConnection, IDictionary<string, object> parameters = null) {
			var processName = GetSyncProcessName(parameters);
			var senderEmailAddress = GetParametersValue<string>(parameters, "SenderEmailAddress");
			var periodInMinutes = GetParametersValue(parameters, "PeriodInMinutes", 0);
			return GetSyncJobName(userConnection, processName, senderEmailAddress, GetSyncJobSuffix(periodInMinutes));
		}

		/// <summary>
		/// Get synchronization process name.
		/// </summary>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <param name="defValue">Default value.</param>
		/// <returns>Synchronization process name.</returns>
		private string GetSyncProcessName(IDictionary<string, object> parameters, string defValue = null) {
			if (!string.IsNullOrEmpty(defValue)) {
				return defValue;
			}
			return GetProcessName(parameters);
		}

		/// <summary>
		/// Get synchronization process name.
		/// </summary>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <returns>Synchronization process name.</returns>
		private string GetProcessName(IDictionary<string, object> parameters) {
			var mailboxType = GetParametersValue<Guid>(parameters, "MailboxType");
			if(mailboxType == Guid.Empty) {
				throw new ArgumentNullException("MailboxType");
			}
			return mailboxType == MailServer.ImapTypeId ? "SyncImapMail" : "LoadExchangeEmailsProcess";
		}

		/// <summary>
		/// Get value of <paramref name="parameters"/> by <paramref name="key"/>, or <paramref name="defValue"/>.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="parameters"></param>
		/// <param name="key"></param>
		/// <param name="defValue">Default parameter value.</param>
		/// <returns>Value of <paramref name="parameters"/></returns>
		private T GetParametersValue<T>(IDictionary<string, object> parameters, string key, T defValue = default) {
			if (!parameters.ContainsKey(key) && defValue == null) {
				throw new ArgumentNullException(key);
			}
			return parameters.ContainsKey(key) ? (T)parameters[key] : defValue;
		}

		/// <summary>
		/// Removes <paramref name="syncJobName"/> job from scheduled jobs.
		/// </summary>
		/// <param name="syncJobName">Synchronization job name.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		private void RemoveProcessJob(string syncJobName, UserConnection userConnection) {
			var stackTrace = new System.Diagnostics.StackTrace(false);
			_log.ErrorFormat("RemoveJob called: CurrentUser {0}, SyncJobName {1}. Trace: {2}",
				userConnection.CurrentUser.Name, syncJobName, stackTrace.ToString());
			_appSchedulerWraper.RemoveJob(syncJobName, SyncJobGroupName);
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
			var syncJobParameters = GetSyncJobParameters(parameters);
			if (period == 0) {
				syncJobParameters.Add("CreateReminding", true);
				_log.ErrorFormat("ScheduleImmediateProcessJob called: CurrentUser {0}, SyncJobName {1}",
					userConnection.CurrentUser.Name, syncJobName);
				_appSchedulerWraper.ScheduleImmediateProcessJob(syncJobName, SyncJobGroupName, processName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, syncJobParameters);
			} else {
				_log.ErrorFormat("ScheduleMinutelyJob called: CurrentUser {0}, SyncJobName {1}",
					userConnection.CurrentUser.Name, syncJobName);
				_appSchedulerWraper.ScheduleMinutelyJob(syncJobName, SyncJobGroupName, processName,
					userConnection.Workspace.Name, userConnection.CurrentUser.Name, period, syncJobParameters);
			}
		}

		/// <summary>
		/// Get parameters for <see cref="IAppSchedulerWraper"/> scheduler.
		/// </summary>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <returns>Parameters for <see cref="IAppSchedulerWraper"/> scheduler.</returns>
		private IDictionary<string, object> GetSyncJobParameters(IDictionary<string, object> parameters) {
			var result = new Dictionary<string, object>();
			foreach (var parametr in parameters) {
				if (_syncJobParameters.Contains(parametr.Key)) {
					result.Add(parametr.Key, parametr.Value);
				}
			}
			return result;
		}

		/// <summary>
		/// Gets synchronization job suffix.
		/// </summary>
		/// <param name="periodInMinutes">Synchronization period in minutes.</param>
		/// <returns>Synchronization job suffix.</returns>
		private string GetSyncJobSuffix(int periodInMinutes) {
			return periodInMinutes == 0 ? _syncJobSuffix : string.Empty;
		}

		/// <summary>
		/// Gets synchronization job waiting priod in minutes.
		/// </summary>
		/// <param name="parameters">Synchronization process parameters.</param>
		/// <param name="periodInMinutes">Default period.</param>
		/// <returns>Synchronization job waiting priod.</returns>
		private int GetSyncPeriod(IDictionary<string, object> parameters, int periodInMinutes = 0) {
			var syncPeriodInMinutes = GetParametersValue(parameters, "PeriodInMinutes", periodInMinutes);
			if (syncPeriodInMinutes < 0) {
				throw new ArgumentOutOfRangeException("PeriodInMinutes");
			}
			return syncPeriodInMinutes;
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="ISyncJobScheduler.CreateSyncJob(UserConnection, int, string, IDictionary{string, object})"/>
		public void CreateSyncJob(UserConnection userConnection, int periodInMinutes, string processName,
				IDictionary<string, object> parameters = null) {
			int syncPeriodInMinutes = GetSyncPeriod(parameters, periodInMinutes);
			var suffix = GetSyncJobSuffix(syncPeriodInMinutes);
			var senderEmailAddress = GetParametersValue(parameters, "SenderEmailAddress", string.Empty);
			var syncProcessName = GetSyncProcessName(parameters, processName);
			var syncJobName = string.IsNullOrEmpty(senderEmailAddress)
				? GetSyncJobName(userConnection, syncProcessName)
				: GetSyncJobName(userConnection, syncProcessName, senderEmailAddress, suffix);
			CreateProcessJob(userConnection, syncJobName, processName, periodInMinutes, parameters);
		}

		/// <inheritdoc cref="ISyncJobScheduler.CreateSyncJob(UserConnection, IDictionary{string, object})"/>
		public void CreateSyncJob(UserConnection userConnection, IDictionary<string, object> parameters) {
			var processName = GetProcessName(parameters);
			int syncPeriodInMinutes = GetSyncPeriod(parameters);
			CreateSyncJob(userConnection, syncPeriodInMinutes, processName, parameters);
		}

		/// <inheritdoc cref="ISyncJobScheduler.RemoveSyncJob(UserConnection, string, string)"/>
		public void RemoveSyncJob(UserConnection userConnection, string senderEmailAddress, string processName) {
			string syncJobName = GetSyncJobName(userConnection, processName, senderEmailAddress);
			RemoveProcessJob(syncJobName, userConnection);
		}

		/// <inheritdoc cref="ISyncJobScheduler.DoesSyncJobExist(UserConnection, string, string, int)"/>
		public bool DoesSyncJobExist(UserConnection userConnection, IDictionary<string, object> parameters = null) {
			var syncJobName = GetSyncJobName(userConnection, parameters);
			return _appSchedulerWraper.DoesJobExist(syncJobName, SyncJobGroupName);
		}

		#endregion

	}

	#endregion

}
