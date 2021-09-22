namespace Terrasoft.Configuration
{
	using Terrasoft.Core;
	using Terrasoft.Core.Configuration;
	using Terrasoft.Core.Factories;

	#region Class: ListenerServiceFailJobFactory

	/// <summary>
	/// Provides methods for creating the exchange events listener service failure processing job.
	/// </summary>
	[DefaultBinding(typeof(IListenerServiceFailJobFactory))]
	public class ListenerServiceFailJobFactory: IListenerServiceFailJobFactory
	{

		#region Fields: Public

		/// <summary>
		/// Scheduler job group name.
		/// </summary>
		public static readonly string JobGroupName = "ExchangeListener";

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IListenerServiceFailJobFactory.ScheduleListenerServiceFailJob"/>
		public void ScheduleListenerServiceFailJob(UserConnection userConnection) {
			var schedulerWraper = ClassFactory.Get<IAppSchedulerWraper>();
			if (schedulerWraper.DoesJobExist(typeof(ListenerServiceFailJob).FullName, JobGroupName)) {
				return;
			}
			schedulerWraper.RemoveGroupJobs(JobGroupName);
			SysUserInfo currentUser = userConnection.CurrentUser;
			int periodMin = Terrasoft.Core.Configuration.SysSettings.GetValue(userConnection, "ListenerServiceFailJobPeriod", 1);
			if (periodMin == 0) {
				return;
			}
			schedulerWraper.ScheduleMinutelyJob<ListenerServiceFailJob>(JobGroupName, userConnection.Workspace.Name,
				currentUser.Name, periodMin, null, true);
		}

		#endregion

	}

	#endregion

}