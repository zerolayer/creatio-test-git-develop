namespace Terrasoft.Sync.Exchange
{
	using System;
	using Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Sync;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeTask

	/// <summary>
	/// Provides methods for tasks synchronization with Exchange.
	/// </summary>
	[Map("Activity", 0, IsPrimarySchema = true, Direction = SyncDirection.DownloadAndUpload, FetchColumnNames = new[] {
			"Title", "StartDate", "DueDate", "Status", "Priority", "RemindToOwner", "RemindToOwnerDate", "Notes" })]
	public class ExchangeTask : ExchangeBase
	{

		#region Constants: Private

		private const string InvalidItemPropertyMessageTemplate =
				"[ExchangeTask.FillLocalItem]: In Exchange task {1} property is empty. Id: {0}";

		private const string InvalidItemTypeMessageTemplate =
				"[ExchangeTask.FillLocalItem]: Invalid Item type. Id: {0}, Subject: {1}";

		#endregion

		#region Fields: Private

		private static Exchange.PropertySet _propertySet =
			new Exchange.PropertySet(Exchange.BasePropertySet.FirstClassProperties);

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Creates new <see cref="ExchangeTask"/> instance, using <paramref name="schema"/> and remote storage item.
		/// </summary>
		/// <param name="schema">Entity sync schema instance.</param>
		/// <param name="item">Remote storage item.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeTask(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timeZoneInfo)
			: base(schema, item, timeZoneInfo) {
			_propertySet.RequestedBodyType = Exchange.BodyType.Text;
			_propertySet.Add(ExchangeUtilityImpl.LocalIdProperty);
		}

		/// <summary>
		/// Creates new  <see cref="ExchangeTask"/> instance, using <paramref name="schema"/>, remote storage item 
		/// instance and remote storage item id.
		/// </summary>
		/// <param name="schema">Entity sync schema instance.</param>
		/// <param name="item">Remote storage item.</param>
		/// <param name="remoteId">Remote storage item id.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeTask(SyncItemSchema schema, Exchange.Item item, string remoteId, TimeZoneInfo timeZoneInfo)
			: this(schema, item, timeZoneInfo) {
			RemoteId = remoteId;
		}

		#endregion

		#region Methods: Private

		private bool CheckTask(SyncContext context, Exchange.Task exchangeTask) {
			if (exchangeTask == null) {
				LogError(context, InvalidItemTypeMessageTemplate, Item.Subject);
				return false;
			}
			if (string.IsNullOrEmpty(exchangeTask.Subject)) {
				LogError(context, InvalidItemPropertyMessageTemplate, "Subject");
				return false;
			}
			return true;
		}
		
		private void SetActivityRemindToOwner(SyncContext context, Entity activity, Exchange.Task exchangeTask) {
			try {
				activity.SetColumnValue("RemindToOwner", exchangeTask.IsReminderSet);
			}
			catch (ServiceObjectPropertyException) {
				activity.SetColumnValue("RemindToOwner", false);
				LogError(context, InvalidItemPropertyMessageTemplate, "IsReminderSet");
			}
		}

		/// <summary>
		/// Sets <see cref="Task.CompleteDate"/> and <see cref="TaskStatus.Completed"/> of exchange task.
		/// </summary>
		/// <param name="exchangeTask"><see cref="Exchange.Task/> instance.</param>
		/// <remarks>If <see cref="Task.CompleteDate"/> is in future, field will be filled on Exchange server.
		/// </remarks>
		/// <remarks>If <see cref="Task.CompleteDate"/> is in the past and <see cref="Task.Status"/> value is set,
		/// than <see cref="Task.CompleteDate"/> will be set as current date. If <see cref="Task.CompleteDate"/> value
		/// set, than <see cref="Task.Status"/> will be set as <see cref="TaskStatus.Completed"/>.
		/// </remarks>
		private void SetRemoteCompleteStatusAndDate(Task exchangeTask) {
			if ((exchangeTask.DueDate != null) && (exchangeTask.DueDate.Value != null)) {
				DateTime dueDate = exchangeTask.DueDate.Value.Date;
				if (dueDate < DateTime.UtcNow) {
					exchangeTask.CompleteDate = dueDate;
					return;
				}
			}
			exchangeTask.Status = TaskStatus.Completed;
		}

		private void LogError(SyncContext context, string message, string propertyName) {
			context.LogError(Action, SyncDirection.Upload, message, Item.Id.UniqueId, propertyName);
		}

		private DateTime? InitExchangeTaskDate (DateTime? date, UserConnection userConnection) {
			return date == null ? (DateTime?)null : (date.Value.GetUserDateTime(userConnection)).Date;
		}
		
		private DateTime? GetExchangeTaskDueDate (Exchange.Task exchangeTask) {
			return (exchangeTask.Status == Exchange.TaskStatus.Completed)
				? exchangeTask.CompleteDate
				: exchangeTask.DueDate;
		}
		
		private void setExchangeTaskExtendedProperty(Exchange.Task exchangeTask, string activityId) {
			exchangeTask.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, activityId);
			exchangeTask.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite);
		}
		
		#endregion

		#region Methods: Protected
		
		/// <summary>
		/// Calls <see cref="Exchange.Item.Load(Exchange.PropertySet)"/> method for <paramref name="exchangeTask"/>.
		/// </summary>
		/// <param name="exchangeTask"><see cref="Exchange.Task"/> instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual void LoadItemProperties(Exchange.Task exchangeTask) {
			exchangeTask.Load(_propertySet);
		}

		/// <summary>
		/// Sets<paramref name="exchangeTask"/> properties to <paramref name="activity"/>.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="exchangeTask"><see cref="Exchange.Task"/> instance.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="action">Current synchronization action.</param>
		protected virtual void SetActivityProperties(Entity activity, Exchange.Task exchangeTask, SyncContext context,
				SyncAction action) {
			activity.SetColumnValue("ShowInScheduler", false);
			DateTime startDate;
			DateTime dueDate;
			SetStartAndDueDate(context.UserConnection, exchangeTask, out startDate, out dueDate);
			if (action == SyncAction.Update) {
				startDate = startDate.Date + activity.GetTypedColumnValue<DateTime>("StartDate").TimeOfDay;
				dueDate = dueDate.Date + activity.GetTypedColumnValue<DateTime>("DueDate").TimeOfDay;
			}
			activity.SetColumnValue("Title", exchangeTask.Subject);
			activity.SetColumnValue("StartDate", startDate);
			activity.SetColumnValue("DueDate", dueDate);
			activity.SetColumnValue("OwnerId", context.UserConnection.CurrentUser.ContactId);
			activity.SetColumnValue("AuthorId", context.UserConnection.CurrentUser.ContactId);
			activity.SetColumnValue("PriorityId", exchangeTask.Importance.GetActivityPriority());
			activity.SetColumnValue("StatusId", exchangeTask.Status.GetActivityStatus());
			SetActivityRemindToOwner(context, activity, exchangeTask);
			if (activity.GetTypedColumnValue<bool>("RemindToOwner")) {
				activity.SetColumnValue("RemindToOwnerDate", exchangeTask.SafeGetValue<DateTime>(
					Exchange.TaskSchema.ReminderDueBy).GetUserDateTime(context.UserConnection));
			} else {
				activity.SetColumnValue("RemindToOwnerDate", null);
			}
			activity.SetColumnValue("Notes", exchangeTask.Body.Text);
		}

		/// <summary>
		/// Checks are the properties of the <paramref name="exchangeTask"/> differ from the properties of 
		/// the <paramref name="activity"/>.
		/// </summary>
		/// <param name="exchangeTask"><see cref="Exchange.Task"/> instance.</param>
		/// <param name="activity">Activity instance.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c> if <paramref name="exchangeTask"/> properties changed. Returns <c>false</c>
		/// otherwise.</returns>
		protected virtual bool IsTaskChanged(Exchange.Task exchangeTask, Entity activity,
				SyncContext context) {
			var userConnection = context.UserConnection;
			string oldActivityHash = GetActivityHash(activity, userConnection);
			EntitySchema activitySchema = userConnection.EntitySchemaManager.GetInstanceByName("Activity");
			Entity newActivity = activitySchema.CreateEntity(userConnection);
			SetActivityProperties(newActivity, exchangeTask, context, SyncAction.Create);
			string activityHash = GetActivityHash(newActivity, userConnection);
			return activityHash != oldActivityHash;
		}

		/// <summary>
		/// Returns unique hash for the <paramref name="activity"/> instance.
		/// <seealso cref="ActivityUtils.GetActivityHash"/>.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Unique hash for the <paramref name="activity"/> instance.</returns>
		protected virtual string GetActivityHash(Entity activity, UserConnection userConnection) {
			TimeZoneInfo currentTimeZone = userConnection.CurrentUser.TimeZone;
			return ActivityUtils.GetActivityHash(activity.GetTypedColumnValue<string>("Title"),
				string.Empty, activity.GetTypedColumnValue<DateTime>("StartDate").Date,
				activity.GetTypedColumnValue<DateTime>("DueDate").Date, activity.GetTypedColumnValue<Guid>("PriorityId"),
				activity.GetTypedColumnValue<Guid>("StatusId").ToString(), currentTimeZone);
		}
		
		#endregion

		#region Methods: Public

		/// <summary>
		/// Sets start and due dates for task.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="startDate">Start date of task.</param>
		/// <param name="dueDate">Due date of task.</param>
		/// <param name="exchangeTask">Exchange task.</param>
		/// <remarks>
		/// When start and due dates are null,start date sets current date and due date sets start date add one day.
		/// When start day is not null and due date is null,due date sets start date add one day.
		/// When due day is not null and start date is null,start date sets due date minus one day.
		/// </remarks>
		public void SetStartAndDueDate(UserConnection userConnection, Exchange.Task exchangeTask,
			out DateTime startDate, out DateTime dueDate) {
			DateTime? exchangeTaskStartDate = InitExchangeTaskDate(exchangeTask.StartDate, userConnection);
			DateTime? exchangeTaskDueDate = InitExchangeTaskDate(GetExchangeTaskDueDate(exchangeTask), userConnection);
			if (exchangeTaskStartDate == null) {
				if (exchangeTaskDueDate == null) {
					startDate = DateTime.Now.Date;
				} else {
					dueDate = exchangeTaskDueDate.Value.Date;
					startDate = dueDate.AddDays(-1);
				}
			} else {
				startDate = exchangeTaskStartDate.Value;
			}
			dueDate = (exchangeTaskDueDate == null)
				? startDate
				: exchangeTaskDueDate.Value.Date;
		}

		/// <summary>
		/// Fills sync item in local storage <paramref name="localItem"/> from sync item in external storage.
		/// </summary>
		/// <param name="localItem">Sync item in local storage.</param>
		/// <param name="context">Sync context.</param>
		public override void FillLocalItem(SyncContext context, ref LocalItem localItem) {
			if (IsDeletedProcessed("Activity", ref localItem)) {
				return;
			}
			var exchangeTask = Item as Exchange.Task;
			if (!CheckTask(context, exchangeTask) || GetRemoteItemLockedForSync(context)) {
				localItem.Entities["Activity"].ForEach(se => se.Action = SyncAction.None);
				Action = SyncAction.None;
				return;
			}
			LoadItemProperties(exchangeTask);
			var activity = GetActivity(context, exchangeTask, ref localItem);
			var action = localItem.Entities["Activity"][0].Action;
			if (action == SyncAction.Create) {
				setExchangeTaskExtendedProperty(exchangeTask, activity.PrimaryColumnValue.ToString());
			}
			if (action == SyncAction.Update && !IsTaskChanged(exchangeTask, activity, context)) {
				LogInfo(context, SyncAction.None, SyncDirection.Download,
					"Task \"{0}\" (Activity Id = {1}) not changed, item skipped.", exchangeTask.Subject,
						activity.PrimaryColumnValue);
				Action = SyncAction.None;
				localItem.Entities["Activity"].ForEach(se => se.Action = SyncAction.None);
				return;
			}
			SetActivityProperties(activity, exchangeTask, context, action);
		}

		/// <summary>
		/// Fills sync item in external storage from sync item in local storage. 
		/// <paramref name="localItem"/>.
		/// </summary>
		/// <param name="localItem">Sync item from local storage.</param>
		/// <param name="context">Sync context.</param>
		public override void FillRemoteItem(SyncContext context, LocalItem localItem) {
			if (localItem.Entities["Activity"][0].State == SyncState.Deleted) {
				Action = SyncAction.Delete;
				return;
			}
			var exchangeTask = (Exchange.Task)Item;
			if (Action == SyncAction.None) {
				return;
			}
			TimeZoneInfo userTimeZone = context.UserConnection.CurrentUser.TimeZone;
			var activity = GetEntityInstance<Entity>(context, localItem, "Activity");
			if (IsOldActivity(activity.GetTypedColumnValue<DateTime>("DueDate"), context) || GetEntityLockedForSync(activity.PrimaryColumnValue, context)){
				Action = SyncAction.None;
				return;
			}
			exchangeTask.Subject = activity.GetTypedColumnValue<string>("Title");
			exchangeTask.StartDate = 
				TimeZoneInfo.ConvertTimeToUtc(activity.GetTypedColumnValue<DateTime>("StartDate"), userTimeZone).ToLocalTime();
			exchangeTask.DueDate = 
				TimeZoneInfo.ConvertTimeToUtc(activity.GetTypedColumnValue<DateTime>("DueDate"), userTimeZone).ToLocalTime();
			exchangeTask.Importance = (Exchange.Importance)ExchangeUtility.GetExchangeImportance(activity.GetTypedColumnValue<Guid>("PriorityId"));
			exchangeTask.IsReminderSet = activity.GetTypedColumnValue<bool>("RemindToOwner");
			if (exchangeTask.IsReminderSet) {
				var remindToOwnerDate = activity.GetTypedColumnValue<DateTime>("RemindToOwnerDate");
				if (remindToOwnerDate != DateTime.MinValue) {
					exchangeTask.ReminderDueBy = TimeZoneInfo.ConvertTimeToUtc(remindToOwnerDate, userTimeZone).ToLocalTime();
				}
			}
			exchangeTask.Body = new MessageBody(BodyType.HTML, activity.GetTypedColumnValue<string>("Notes"));
			TaskStatus exchangeStatus = (TaskStatus)ExchangeUtility.GetExchangeTaskStatus(activity.GetTypedColumnValue<Guid>("StatusId"));
			if (exchangeStatus == TaskStatus.Completed) {
				SetRemoteCompleteStatusAndDate(exchangeTask);
			} else {
				exchangeTask.Status = exchangeStatus;
			}
			exchangeTask.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, activity.PrimaryColumnValue.ToString());
		}

		#endregion
	}

	#endregion
}