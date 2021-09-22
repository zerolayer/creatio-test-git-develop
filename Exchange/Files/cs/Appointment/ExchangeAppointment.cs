namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Data;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using Newtonsoft.Json.Linq;
	using Terrasoft.Sync;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Common;
	using Terrasoft.Common.Json;
	using Terrasoft.Core.Entities;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeAppointment

	/// <summary>
	/// Represents a class object synchronization with Exchange appointment repository.
	/// </summary>
	[Map("Activity", 0, IsPrimarySchema = true, Direction = SyncDirection.DownloadAndUpload, FetchColumnNames = new[] {
		"Title", "StartDate", "DueDate", "Priority", "RemindToOwner", "RemindToOwnerDate", "Notes", "Status",
		"Location"})]
	[Map("ActivityParticipant", 1, PrimarySchemaName = "Activity", ForeingKeyColumnName = "Activity",
			Direction = SyncDirection.DownloadAndUpload, FetchColumnNames
				= new[] { "Participant", "InviteResponse", "Role" })]
	public class ExchangeAppointment : ExchangeBase
	{
		#region Constants: Private

		private const int exchangeMaxTitleLength = 255;

		/// <summary>
		/// "Russian Standard Time" time zone identifier. 
		/// </summary>
		private const string RussianStandardTimeTimeZoneId = "Russian Standard Time";

		/// <summary>
		/// Custom "Russian Standard Time" time zone identifier. 
		/// </summary>
		private const string CustomRussianStandardTimeTimeZoneId = "Custom Russian Standard Time";

		#endregion

		#region Fields: Private

		private static Exchange.PropertySet _propertySet =
			new Exchange.PropertySet(Exchange.BasePropertySet.FirstClassProperties);

		/// <summary>
		/// The object of activity synchronization.
		/// </summary>
		private SyncEntity _activitySyncEntity;

		/// <summary>
		/// Activity participant identifiers.
		/// </summary>
		private Dictionary<string, Guid> _participantsRoles;

		#endregion

		#region Constructors: Public

		/// <summary>Initializes a new instance of the class <see cref="ExchangeAppointment"/>,
		/// using the specified schema and a basic element in the external storage.</summary>
		/// <param name="schema">Scheme sync element.</param>
		/// <param name="item">Basic element in the external storage.</param>
		/// <param name="timeZoneInfo">Local time.</param>
		public ExchangeAppointment(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timeZoneInfo)
			: base(schema, item, timeZoneInfo) {
			_propertySet.RequestedBodyType = Exchange.BodyType.HTML;
			_propertySet.Add(ExchangeUtilityImpl.LocalIdProperty);
		}

		/// <summary>
		/// Initializes a new instance of the class <see cref="ExchangeAppointment"/>,
		/// using the specified schema , a basic element in the external storage and
		/// <paramref name="remoteId"/> with local time.
		/// </summary>
		/// <param name="schema">Scheme sync element.</param>
		/// <param name="item">Basic element in the external storage.</param>
		/// <param name="remoteId">Unique identifier element in the external storage.</param>
		/// <param name="timeZoneInfo">Local time.</param>
		public ExchangeAppointment(SyncItemSchema schema, Exchange.Item item, string remoteId,
			TimeZoneInfo timeZoneInfo)
			: this(schema, item, timeZoneInfo) {
			RemoteId = remoteId;
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// Unique item id in remote storage.
		/// </summary>
		public override string Id {
			get {
				if (RemoteId.IsNullOrEmpty()) {
					var item = (Exchange.Appointment)Item;
					RemoteId = item.ICalUid;
					if (item.SafeGetValue<bool>(Exchange.AppointmentSchema.IsRecurring) && item.ICalRecurrenceId != null) {
						RemoteId += item.ICalRecurrenceId.Value.ToString("_yyyy_MM_dd");
					}
				}
				return RemoteId;
			}
			internal set {
				RemoteId = value;
			}
		}

		/// <summary>
		/// Flag sending Exchange notification.
		/// </summary>
		public bool SendNotifications {
			get;
			set;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Verifies the appointment item from external storage.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		/// <returns>Flag indicating if the appointment is not empty.</returns>
		private bool CheckAppointment(SyncContext context, Exchange.Appointment exchangeAppointment) {
			if (exchangeAppointment == null) {
				LogError(context, Action, SyncDirection.Upload,
					"[ExchangeAppointment.FillLocalItem]: Invalid Item type. Id: {0}, Subject: {1}", GetItemSimpleId(),
					Item.Subject);
				return false;
			}
			if (string.IsNullOrEmpty(exchangeAppointment.ICalUid)) {
				LogError(context, Action, SyncDirection.Upload,
					"[ExchangeAppointment.FillLocalItem]: In Exchange appointment has no ICalUid (Probably deleted from Exchange). Id: {0}",
					GetItemSimpleId());
				return false;
			}
			return true;
		}

		/// <summary>
		/// Sets SendNotification parameter.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="activity">Exchange appointment item in local storage.</param>
		/// <param name="syncEntity">The object of activity synchronization.</param>
		/// </returns>
		private void SetSendNotification(SyncContext context, Entity activity, SyncEntity syncEntity) {
			SendNotifications = true;
			if (!((string)syncEntity.ExtraParameters).IsNullOrEmpty()) {
				JObject activityExtendProperty = JObject.Parse(syncEntity.ExtraParameters);
				Guid previousStatusId = Guid.Parse((string)activityExtendProperty["StatusId"]);
				DateTime previousDueDate = DateTime.Parse((string)activityExtendProperty["EndDate"]);
				SendNotifications = IsActivityStatusTypeChanged(activity, previousStatusId, context.UserConnection) &&
						IsActivityDueDateChanged(activity, previousDueDate, context);
			} else {
				SendNotifications = !IsActivityStatusFinal(activity.GetTypedColumnValue<Guid>("StatusId"), context.UserConnection) &&
					!IsActivityDueDateInPast(activity, context);
			}
		}

		///<summary>Checks if activity status is final.</summary>
		/// <param name="activityStatus">Activity status id.</param>
		/// <param name="userConnection">User connection instance.</param>
		/// <returns>Is activity status in final state.</returns>
		private bool IsActivityStatusFinal(Guid activityStatus, UserConnection userConnection) {
			Select select = new Select(userConnection).From("ActivityStatus")
				.Column("Finish")
				.Where("Id").IsEqual(Column.Parameter(activityStatus)) as Select;
			return select.ExecuteScalar<bool>();
		}

		/// <summary>Checks if activity due date in past.</summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="context">Synchronization context.</param>
		/// <returns>Is activity due date in past.</returns>
		private bool IsActivityDueDateInPast(Entity activity, SyncContext context) {
			DateTime currentDateTime = context.CurrentSyncStartVersion.ToLocalTime();
			return (activity.GetTypedColumnValue<DateTime>("DueDate").ToLocalTime() < currentDateTime);
		}

		/// <summary>
		/// Sets activity title when length more than 255 symbols.
		/// </summary>
		/// <param name="activity"> Instance of Activity.</param>
		/// <param name="exchangeSubject">Appointment subject.</param>
		private void SetActivityTitle(Entity activity, string exchangeSubject) {
			var exchangeTitleLength = exchangeSubject.Length - 3;
			string rawTitle = activity.GetIsColumnValueLoaded("Title")
				? activity.GetTypedColumnValue<string>("Title")
				: string.Empty;
			string resultTitle = exchangeSubject;
			if (exchangeSubject.Length == exchangeMaxTitleLength && 
				exchangeSubject.Substring(exchangeTitleLength) == "..." && rawTitle.Length > exchangeMaxTitleLength) {
				resultTitle = exchangeSubject.Substring(0, exchangeTitleLength) + 
					rawTitle.Substring(exchangeTitleLength);
			}
			activity.SetColumnValue("Title", resultTitle);
		}

		/// <summary>
		/// Checks if activity status not changed from final state to final state.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="previousStatusId">Activity previous state id.</param>
		/// <param name="userConnection">User connection instance.</param>
		/// <returns>Is activity status not changed from final state to final state.</returns>
		private bool IsActivityStatusTypeChanged(Entity activity, Guid previousStatusId,
				UserConnection userConnection) {
			return !IsActivityStatusFinal(activity.GetTypedColumnValue<Guid>("StatusId"), userConnection) &&
					!IsActivityStatusFinal(previousStatusId, userConnection);
		}

		/// <summary>
		/// Checks if activity due date not changed from past to past.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="previousDueDate">Activity previous due date.</param>
		/// <param name="context">Synchronization context.</param>
		/// <returns>Is activity due date not changed from past to past.</returns>
		private bool IsActivityDueDateChanged(Entity activity, DateTime previousDueDate, SyncContext context) {
			DateTime currentDateTime = context.CurrentSyncStartVersion.ToLocalTime();
			return (activity.GetTypedColumnValue<DateTime>("DueDate").ToLocalTime() >= currentDateTime) ||
					(previousDueDate.ToLocalTime() >= currentDateTime);
		}

		/// <summary>
		/// Creates <see cref="Select"/> for entity containing contacts email address.
		/// </summary>
		/// <apram name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <apram name="entityName">Entity containing contacts email address.</param>
		/// <apram name="contactIdColumn">Column containing contact id.</param>
		/// <apram name="emailColumn">Column containing contact email.</param>
		/// <returns>
		/// <see cref="Select"/> instance.
		/// </returns>
		private Select GetContactEmailsSelect(UserConnection userConnection, string entityName, string contactIdColumn,
				string emailColumn) {
			return new Select(userConnection)
					.Column(emailColumn)
					.Column(contactIdColumn)
					.From(entityName);
		}

		/// <summary>
		/// Gets a list of email addresses participants.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="activityId">The unique identifier of the activity.</param>
		/// <param name="rolesId">Roles for selection of participant.</param>
		/// <returns>List of email addresses participants.
		/// </returns>
		private List<string> GetMeetingParticipantEmails(UserConnection userConnection, Guid activityId, List<Guid> rolesId) {
			var result = new List<string>();
			bool participantsExists = false;
			Guid userContactId = userConnection.CurrentUser.ContactId;
			var selectParticipants = new Select(userConnection).Column("ParticipantId")
					.From("ActivityParticipant")
					.Where("ActivityId").IsEqual(Column.Parameter(activityId))
					.And("RoleId").In(Column.Parameters(rolesId));
			Select selectAttendeeEmails;
			string contactIdColumnName;
			string emailColumnName;
			if (userConnection.GetIsFeatureEnabled("UseAllEmailsForExchangeAppointments")) {
				contactIdColumnName = "ContactId";
				emailColumnName = "Number";
				selectAttendeeEmails = GetContactEmailsSelect(userConnection,
						"ContactCommunication", contactIdColumnName, emailColumnName)
						.Where("CommunicationTypeId").IsEqual(Column.Parameter(new Guid(CommunicationTypeConsts.EmailId)))
						.And("ContactId").In(selectParticipants) as Select;
			} else {
				contactIdColumnName = "Id";
				emailColumnName = "Email";
				selectAttendeeEmails = GetContactEmailsSelect(userConnection,
						"Contact", contactIdColumnName, emailColumnName)
						.Where("Email").IsNotEqual(Func.IsNull(Column.Const(string.Empty), Column.Const("null")))
						.And("Id").In(selectParticipants) as Select;
			}
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = selectAttendeeEmails.ExecuteReader(dbExecutor)) {
					while (reader.Read()) {
						var attendeeEmail = reader.GetColumnValue<string>(emailColumnName);
						if (!string.IsNullOrEmpty(attendeeEmail)) {
							Guid contactId = reader.GetColumnValue<Guid>(contactIdColumnName);
							if (!contactId.Equals(userContactId)) {
								participantsExists = true;
							}
							result.Add(attendeeEmail);
						}
					}
				}
			}
			if (participantsExists) {
				return result;
			}
			return new List<string>();
		}

		/// <summary>Returns sync action for activity participant instance.</summary>
		/// <param name="participant">Activity participant instance.</param>
		/// <param name="attendee"><see cref="Exchange.Attendee"/> instance.</param>
		/// <returns>Sync action for activity participant.</returns>
		private SyncAction GetParticipantSyncAction(Entity participant, Exchange.Attendee attendee) {
			Guid inviteResponseId = GetParticipantInviteResponse(attendee);
			if (participant.StoringState == Terrasoft.Core.Configuration.StoringObjectState.New) {
				return SyncAction.Create;
			}
			if (!inviteResponseId.Equals(participant.GetTypedColumnValue<Guid>("InviteResponseId"))) {
				return SyncAction.Update;
			}
			return SyncAction.None;
		}

		/// <summary>
		/// Returns <see cref="ParticipantResponse"/> id for <paramref name="attendee"/> invite response.
		/// </summary>
		/// <param name="attendee"><see cref="Exchange.Attendee"/> instance.</param>
		/// <returns>Participant response id.</returns>
		private Guid GetParticipantInviteResponse(Exchange.Attendee attendee) {
			switch (attendee.ResponseType) {
				case Exchange.MeetingResponseType.Accept:
				case Exchange.MeetingResponseType.Organizer:
					return ActivityConsts.ParticipantResponseConfirmedId;
				case Exchange.MeetingResponseType.Decline:
					return ActivityConsts.ParticipantResponseDeclinedId;
				case Exchange.MeetingResponseType.Unknown:
				case Exchange.MeetingResponseType.Tentative:
					return ActivityConsts.ParticipantResponseInDoubtId;
				default:
					return Guid.Empty;
			}
		}

		/// <summary>Returns activity participant instance.</summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="attendee"><see cref="Exchange.Attendee"/> instance.</param>
		/// <param name="contactId">Contact unique identifier.</param>
		/// <param name="activityId">The unique identifier of the activity.</param>
		/// <param name="localItem">The local storage item.</param>
		/// <param name="roleId">The value to filling role of <see cref="ActivityParticipant"/>.</param>
		/// <returns>Activity participant instance.</returns>
		private Entity GetParticipant(UserConnection userConnection, Exchange.Attendee attendee,
				Guid contactId, Guid activityId, LocalItem localItem, Guid roleId) {
			Select isParticipantExists = new Select(userConnection)
					.Column("Id")
					.From("ActivityParticipant")
				.Where("ActivityId").IsEqual(Column.Parameter(activityId))
				.And("ParticipantId").IsEqual(Column.Parameter(contactId)) as Select;
			Guid existingParticipantId = isParticipantExists.ExecuteScalar<Guid>();
			var schema = userConnection.EntitySchemaManager.GetInstanceByName("ActivityParticipant");
			Entity participant = schema.CreateEntity(userConnection);
			if (existingParticipantId.IsEmpty() || !participant.FetchFromDB(existingParticipantId, false)) {
				participant.SetDefColumnValues();
				participant.SetColumnValue("ParticipantId", contactId);
				participant.SetColumnValue("ActivityId", activityId);
			}
			if (participant.GetTypedColumnValue<Guid>("RoleId") != _participantsRoles["Responsible"]) {
				participant.SetColumnValue("RoleId", roleId);
			}
			var syncInstance = SyncEntity.CreateNew(participant);
			syncInstance.Action = GetParticipantSyncAction(participant, attendee);
			localItem.AddOrReplace("ActivityParticipant", syncInstance);
			return participant;
		}

		/// <summary>Syncronizes appointment attendee.</summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="activityId">Activity unique identifier.</param>
		/// <param name="attendee">Exchange appointment attendee.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <param name="roleId">The value to filling role of <see cref="ActivityParticipant"/>.</param>
		/// <param name="organizerEmail">Appointment organizer email.</param>
		/// <returns>List of activity participants unique identifiers.</returns>
		private List<Guid> SyncAttendee(SyncContext context, Guid activityId, Exchange.Attendee attendee, LocalItem localItem, Guid roleId, string organizerEmail) {
			var userConnection = context.UserConnection;
			List<Guid> participantsIds = new List<Guid>();
			List<string> emails = new List<string>();
			string attendeeMail = attendee.AttendeeToEmailAddressString();
			emails.Add(attendeeMail.ExtractEmailAddress());
			List<Guid> contactIds = ContactUtilities.FindContactsByEmail(organizerEmail.ExtractEmailAddress(), userConnection);
			LogInfo(context, "SyncAttendee for \"{0}\" emails started.", string.Join(", ", emails));
			foreach (var contact in ContactUtilities.GetContactsByEmails(userConnection, emails)) {
				Guid contactId = contact.Key;
				Entity participant = GetParticipant(userConnection, attendee, contactId, activityId, localItem, roleId);
				participantsIds.Add(participant.PrimaryColumnValue);
				Guid inviteResponseId = contactIds.Any(cid => cid.Equals(contactId)) ? ActivityConsts.ParticipantResponseConfirmedId : GetParticipantInviteResponse(attendee);
				if (inviteResponseId.IsNotEmpty()) {
					participant.SetColumnValue("InviteResponseId", inviteResponseId);
				}
			}
			return participantsIds;
		}

		/// <summary>
		/// Updates sync action for deleted activity participants.
		/// </summary>
		/// <param name="modifiedParticipants">Modified participants unique identifiers.</param>
		/// <param name="localItem">Local storage item.</param>
		private void ActualizeSyncParticipantsList(List<Guid> modifiedParticipants, LocalItem localItem) {
			List<SyncEntity> entities = localItem.Entities["ActivityParticipant"];
			foreach (SyncEntity entity in entities.Where(e => !modifiedParticipants.Any(p => e.EntityId == p))) {
				ActivityParticipant participant = (ActivityParticipant)entity.Entity;
				if (participant.StoringState == Terrasoft.Core.Configuration.StoringObjectState.Deleted) {
					entity.Action = SyncAction.Delete;
				}
			}
		}

		/// <summary>Syncronizes appointment required attendees.</summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="activity">Activity entity instance.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		/// <param name="localItem">Local storage item.</param>
		private void SyncAllAttendees(SyncContext context, Entity activity, Exchange.Appointment exchangeAppointment,
				LocalItem localItem) {
			_participantsRoles = _participantsRoles ?? ActivityUtils.GetParticipantsRoles(context.UserConnection);
			List<Guid> modifiedParticipants = new List<Guid>();
			var organizerEmail = exchangeAppointment.Organizer.AttendeeToEmailAddressString();
			foreach (Exchange.Attendee attendee in exchangeAppointment.RequiredAttendees) {
				modifiedParticipants.AddRange(SyncAttendee(context, activity.PrimaryColumnValue, attendee, localItem,
					_participantsRoles["Participant"], organizerEmail));
			}
			foreach (Exchange.Attendee attendee in exchangeAppointment.OptionalAttendees) {
				modifiedParticipants.AddRange(SyncAttendee(context, activity.PrimaryColumnValue, attendee, localItem,
					_participantsRoles["OptionalParticipant"], organizerEmail));
			}
			ActualizeSyncParticipantsList(modifiedParticipants, localItem);
		}

		/// <summary>Sets reccuring master local storage item properties.</summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem">Local storage item.</param>
		private void FillReccuringMasterLocalItem(SyncContext context, LocalItem localItem) {
			LogInfo(context, "Appointment {0} starting FillReccuringMasterLocalItem (Id = {1}, action = {2}).", GetDisplayName(), Id, Action);
			if (Action == SyncAction.CreateRecurringMaster && context.UserConnection.GetIsFeatureEnabled("ExchangeRecurringAppointments")) {
				LogInfo(context, "Appointment {0} type is reccuring master (Id = {1}), previous single task deleting.", GetDisplayName(), Id);
				localItem.Entities.ForEach(itemsCollection => itemsCollection.Value.ForEach(item => item.Action = SyncAction.CreateRecurringMaster));
			}
		}

		/// <summary>
		/// Returns current item converted to <see cref="Microsoft.Exchange.WebServices.Data.Appointment"/>.
		/// </summary>
		/// <returns><see cref="Microsoft.Exchange.WebServices.Data.Appointment"/> instance.</returns>
		private Exchange.Appointment GetAppointment() {
			return Item as Exchange.Appointment;
		}

		/// <summary>
		/// Returns true if local item contains Activity instance.
		/// </summary>
		/// <param name="localItem">Local storage item.</param>
		/// <returns>
		/// True if local item contains Activity instance.
		/// </returns>
		private bool IsAppointmentSyncedInBpm(LocalItem localItem) {
			var collection = localItem.Entities["Activity"];
			if (collection.Count == 0) {
				return false;
			}
			var activitySyncEntity = collection[0];
			return (activitySyncEntity.Action != SyncAction.Create);
		}

		/// <summary>
		/// Checks is deleted appointment synchronized.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> when syncronizing not deleted appointment. Otherwise returns <c>false</c>.</returns>
		private bool CheckNotDeleted(SyncContext context, LocalItem localItem) {
			if (IsDeletedProcessed("Activity", ref localItem) ||
					localItem.Entities["Activity"].Any(se => se.State == SyncState.Deleted)) {
				LogInfo(context, "Appointment {0} deleted in bpmonline, fill local item skipped.", GetDisplayName());
				return false;
			}
			return true;
		}

		/// <summary>
		/// Checks is type and ICalUId appointment properties valid.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c> when type and ICalUId appointment properties valid.
		/// Otherwise returns <c>false</c>.</returns>
		private bool CheckTypeAndICalUId(SyncContext context) {
			var exchangeAppointment = GetAppointment();
			if (!CheckAppointment(context, exchangeAppointment)) {
				Action = SyncAction.Delete;
				State = SyncState.Deleted;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Locks item for synchronization to bpmonline.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c> when current process locked appointment for synchronization.
		/// Otherwise returns <c>false</c>.</returns>
		private bool LockItem(SyncContext context) {
			if (GetRemoteItemLockedForSync(context)) {
				LogInfo(context, "Appointment {0} locked for sync in bpmonline (Id = {1}), sync action skipped.", GetDisplayName(), Id);
				Action = SyncAction.None;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Checks appointment organizer properties.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> when current user can be appointment organizer in bpmonline.
		/// Otherwise returns <c>false</c>.</returns>
		private bool CheckOrganizer(SyncContext context, LocalItem localItem) {
			var exchangeAppointment = GetAppointment();
			LoadItemProperties(exchangeAppointment);
			string remoteUId = exchangeAppointment.ICalUid;
			string organizerEmail = GetAppointmentOrganizerEmail(exchangeAppointment);
			LogInfo(context, "Appointment {0} organizerEmail = {1}", GetDisplayName(), organizerEmail);
			if (!GetCanSaveChangesToBpm(context.UserConnection, organizerEmail)) {
				LogInfo(context, "Appointment {0} has organizer in bpmonline (organizer = {1}), sync action skipped.", GetDisplayName(),
					organizerEmail);
				SetLocalItemSchemasAction(context, localItem, SyncAction.None);
				Action = SyncAction.None;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Checks is activity with save content already exists.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> when appointment unique and can be created in bpmonline.
		/// Otherwise returns <c>false</c>.</returns>
		private bool CheckAppointmentContentDuplicates(SyncContext context, LocalItem localItem) {
			var userConnection = context.UserConnection;
			var exchangeAppointment = GetAppointment();
			Entity activity = context.UserConnection.GetIsFeatureEnabled("ExchangeRecurringAppointments")
				? GetEntityInstance<Activity>(context, localItem, "Activity")
				: GetActivity(context, exchangeAppointment, ref localItem);
			if (userConnection.GetIsFeatureEnabled("CheckAppointmentDuplicatesByContent")
					&& IsCreatingDuplicateAppointment(exchangeAppointment, localItem, context)) {
				DateTime startDate = exchangeAppointment.SafeGetValue<DateTime>(Exchange.AppointmentSchema.Start)
					.GetUserDateTime(userConnection);
				DateTime dueDate = exchangeAppointment.SafeGetValue<DateTime>(Exchange.AppointmentSchema.End)
					.GetUserDateTime(userConnection);
				LogInfo(context, SyncAction.None, SyncDirection.Download,
					"Appointment \"{0}\" ({1} - {2}) already exists, item skipped.", exchangeAppointment.Subject,
					startDate, dueDate);
				SetLocalItemSchemasAction(context, localItem, SyncAction.None);
				Action = SyncAction.None;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Checks is current activity instance can be modified in this synchronization session.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> if activity instance can be modified in this synchronization session.
		/// Returns <c>false</c> otherwise.</returns>
		private bool CanModifyActivity(SyncContext context, LocalItem localItem) {
			var activitySyncEntity = localItem.Entities["Activity"][0];
			if (activitySyncEntity.Action != SyncAction.Create) {
				return true;
			}
			if (GetRemoteItemLockedForSync(context, activitySyncEntity.Action)) {
				LogInfo(context, "Appointment {0} locked for sync in bpmonline (Id = {1}), sync action skipped.", GetDisplayName(), Id);
				Action = SyncAction.None;
				SetLocalItemSchemasAction(context, localItem, SyncAction.None);
				return false;
			}
			var exchangeAppointment = GetAppointment();
			Entity activity = activitySyncEntity.Entity;
			if (!SetAppointmentExtendProperty(context, exchangeAppointment, activity)) {
				LogInfo(context, "Appointment {0} marked as new, but has valid activity id in exchange (activity id = {1}), sync action skipped.",
					GetDisplayName(), activity.PrimaryColumnValue);
				SetLocalItemSchemasAction(context, localItem, SyncAction.None);
				Action = SyncAction.None;
				return false;
			}
			return true;
		}

		/// <summary>
		/// Sets activity properties that need for new activites only.
		/// </summary>
		/// <param name="activitySyncEntity"><see cref="SyncEntity"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		private void SetNewActivityProperties(SyncEntity activitySyncEntity, UserConnection userConnection) {
			if (activitySyncEntity.Action == SyncAction.Create) {
				Guid currentContactId = userConnection.CurrentUser.ContactId;
				var activity = activitySyncEntity.Entity;
				activity.SetColumnValue("OwnerId", currentContactId);
				activity.SetColumnValue("StatusId", Guid.Empty);
			}
		}

		/// <summary>
		/// Returns mailbox email address collection.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Mailbox email address collection.</returns>
		private IEnumerable<string> GetCurrentUserMailboxEmails(UserConnection userConnection) {
			var select = new Select(userConnection)
					.Column("mss", "SenderEmailAddress")
					.Column("msd", "Domain")
				.From("MailboxSyncSettings").As("mss")
				.LeftOuterJoin("MailServerDomain").As("msd").On("mss", "MailServerId").IsEqual("msd", "MailServerId")
				.Where("mss", "SysAdminUnitId").IsEqual(Column.Parameter(userConnection.CurrentUser.Id)) as Select;
			var result = new List<string>();
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
					while (reader.Read()) {
						var email = reader.GetColumnValue<string>("SenderEmailAddress");
						var domain = reader.GetColumnValue<string>("Domain");
						result.AddIfNotExists(email.ToLower());
						if (domain.IsNotNullOrEmpty()) {
							var emailWithDomain = string.Concat(email.Split('@')[0], "@", domain);
							result.AddIfNotExists(emailWithDomain.ToLower());
						}
					}
				}
			}
			return result;
		}

		/// <summary>
		/// Gets exchange appointment <see cref="TimeZoneInfo"/>.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>Exchange appointment <see cref="TimeZoneInfo"/>.</returns>
		private TimeZoneInfo GetExchangeAppointmentTimeZoneInfo(UserConnection userConnection) {
			var timeZoneInfo = userConnection.CurrentUser.TimeZone.Id == RussianStandardTimeTimeZoneId
							? TimeZoneInfo.CreateCustomTimeZone(
								CustomRussianStandardTimeTimeZoneId,
								new TimeSpan(3, 0, 0),
								CustomRussianStandardTimeTimeZoneId,
								CustomRussianStandardTimeTimeZoneId)
							: userConnection.CurrentUser.TimeZone;
			var adjustmentRules = timeZoneInfo.GetAdjustmentRules();
			if (adjustmentRules.Any() && !adjustmentRules.Any(ar => ar.DateEnd == DateTime.MaxValue.Date)) {
				var tz = timeZoneInfo;
				return TimeZoneInfo.CreateCustomTimeZone(tz.Id, tz.BaseUtcOffset, tz.DisplayName, tz.StandardName);
			}
			return timeZoneInfo;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Writes log message.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="action"><see cref="SyncAction"/> enumeration.</param>
		/// <param name="syncDirection"><see cref="SyncDirection"/> enumeration.</param>
		/// <param name="format">String a composite format.</param>
		/// <param name="args">An array of additional parameters.</param>
		protected virtual void LogError(SyncContext context, SyncAction action, SyncDirection syncDirection, 
				string format, params object[] args) {
			context.LogError(action, syncDirection, format, args);
		}

		/// <summary>
		/// Returns exchange item unique id.
		/// </summary>
		/// <returns>
		/// Exchange item unique id.
		/// </returns>
		protected virtual string GetItemSimpleId() {
			return Item.Id.UniqueId;
		}

		/// <summary>
		/// Returns unique hash for activity instance.
		/// </summary>
		/// <param name="title">Activity title.</param>
		/// <param name="location">The location of the activity.</param>
		/// <param name="startDate">The start date of the activity.</param>
		/// <param name="dueDate">The due date of the activity.</param>
		/// <param name="priorityId">Unique identifier priority.</param>
		/// <param name="notes">The notes of the activity.</param>
		/// <param name="timeZoneInfo">User timezone.</param>
		/// <returns>Unique hash for activity instance.</returns>
		protected virtual string GetActivityHash(string title, string location, DateTime startDate,
			DateTime dueDate, Guid priorityId, string notes, TimeZoneInfo timeZoneInfo) {
			return ActivityUtils.GetActivityHash(title, location, startDate,
				dueDate, priorityId, notes, timeZoneInfo);
		}

		/// <summary>
		/// Returns <see cref="Exchange.Appointment.Organizer"/> email for <paramref name="exchangeAppointment"/>.
		/// </summary>
		/// <param name="exchangeAppointment"><see cref="Exchange.Appointment"/> instance.</param>
		/// <returns>
		/// Exchange appointment organizer email.
		/// </returns>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual string GetAppointmentOrganizerEmail(Exchange.Appointment exchangeAppointment) {
			return (exchangeAppointment.Organizer.AttendeeToEmailAddressString()).ExtractEmailAddress();
		}

		/// <summary>
		/// Checks if Activity instance with values from <paramref name="appointment"/> exists.
		/// </summary>
		/// <param name="appointment"><see cref="Exchange.Appointment"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <returns><c>True</c>, if Activity instance exists.</returns>
		/// <remarks>
		/// Columns Title, Location, StartDate, DueDate, Priority, Notes used for search. User rights ignored.
		/// </remarks>
		protected virtual bool IsCreatingDuplicateAppointment(Exchange.Appointment appointment, LocalItem localItem,
				SyncContext context) {
			List<SyncEntity> activitySyncEntities = localItem.Entities["Activity"];
			if (activitySyncEntities.IsNotEmpty() && activitySyncEntities.All(
					syncEntity => syncEntity.Action != SyncAction.Create)) {
				return false;
			}
			UserConnection userConnection = context.UserConnection;
			var activitySchema = userConnection.EntitySchemaManager.GetInstanceByName("Activity");
			Activity activity = (Activity)activitySchema.CreateEntity(userConnection);
			activity.StatusId = Guid.Empty;
			SetActivityProporties(userConnection, activity, appointment);
			DateTime utcStartDate = DateTime.SpecifyKind(activity.StartDate, DateTimeKind.Unspecified);
			utcStartDate = TimeZoneInfo.ConvertTimeToUtc(utcStartDate, userConnection.CurrentUser.TimeZone);
			DateTime utcDueDate = DateTime.SpecifyKind(activity.DueDate, DateTimeKind.Unspecified);
			utcDueDate = TimeZoneInfo.ConvertTimeToUtc(utcDueDate, userConnection.CurrentUser.TimeZone);
			Select activitySelect = new Select(userConnection)
					.Column(Func.Count("Id"))
				.From("Activity")
				.Where("Title").IsEqual(Column.Parameter(activity.Title))
					.And("Location").IsEqual(Column.Parameter(activity.Location))
					.And("StartDate").IsEqual(Column.Parameter(utcStartDate))
					.And("DueDate").IsEqual(Column.Parameter(utcDueDate))
					.And("PriorityId").IsEqual(Column.Parameter(activity.PriorityId)) as Select;
			return activitySelect.ExecuteScalar<int>() > 0;
		}

		/// <summary>
		/// Checks that exchange appointment synchronization exists for <paramref name="email"/> account.
		/// <see cref="ActivitySyncSettings.ImportAppointments"/> column value used as active appointment synchronization flag.
		/// <see cref="MailboxSyncSettings.SenderEmailAddress"/> column value used as synchronization settings filtration.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="email">Sender email value.</param>
		/// <returns><c>True</c> if exchange appointment synchronization exists for <paramref name="email"/> account exists, 
		/// <c>false</c> otherwise.</returns>
		protected bool GetAppointmentSynchronizationExists(UserConnection userConnection, string email) {
			if (email.IsNullOrEmpty() || !email.Contains("@")) {
				return false;
			}
			Select settingsSelect = new Select(userConnection).Top(1)
					.Column("MailboxSyncSettings", "Id")
				.From("MailboxSyncSettings")
				.InnerJoin("ActivitySyncSettings").On("MailboxSyncSettings", "Id")
					.IsEqual("ActivitySyncSettings", "MailboxSyncSettingsId")
				.Where("MailboxSyncSettings", "SenderEmailAddress").StartsWith(Column.Parameter(email.Split('@')[0]))
					.And("ActivitySyncSettings", "ImportAppointments").IsEqual(Column.Parameter(true)) as Select;
			return settingsSelect.ExecuteScalar<Guid>().IsNotEmpty();
		}

		/// <summary>
		/// Checks may current user save <see cref="ExchangeBase.Item"/> changes to bpm'online.
		/// Current user may save external item changes to bpm'online if he is appointment organizer, 
		/// or when appointment organizer appointment synchrodization not exists.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="organizerEmail">Appointment organizer email.</param>
		/// <returns><c>True</c> if current user can save <see cref="ExchangeBase.Item"/> changes to bpm'online, 
		/// <c>false</c> otherwise.</returns>
		protected virtual bool GetCanSaveChangesToBpm(UserConnection userConnection, string organizerEmail) {
			Guid currentContactId = userConnection.CurrentUser.ContactId;
			bool currentContactOwner = GetIsContactEmailExist(userConnection, currentContactId, organizerEmail);
			bool ownerHasSynchronization = GetAppointmentSynchronizationExists(userConnection, organizerEmail);
			return currentContactOwner || (!currentContactOwner && !ownerHasSynchronization);
		}

		/// <summary>
		/// Clears attendees if <see cref="SyncAction"/> equal <see cref="SyncAction.Update"/>
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		protected void ClearAttendees(UserConnection userConnection, Exchange.Appointment exchangeAppointment) {
			if (Action != SyncAction.Update) {
				return;
			}
			string organizerEmail = exchangeAppointment.Organizer.AttendeeToEmailAddressString();
			List<Guid> contactIds = ContactUtilities.FindContactsByEmail(organizerEmail.ExtractEmailAddress(), userConnection);
			if (!contactIds.Any(c => c == userConnection.CurrentUser.ContactId)) {
				return;
			}
			exchangeAppointment.RequiredAttendees.Clear();
			exchangeAppointment.OptionalAttendees.Clear();
		}

		/// <summary>
		/// Add appointment participants to <see cref="Exchange.AttendeeCollection"/>.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="activity">Activity instance.</param>
		/// <param name="attendees">Attedees collection.</param>
		/// <param name="roles">Participants roles list.</param>
		protected void AddAppointmentParticipants(UserConnection userConnection, Entity activity, Exchange.AttendeeCollection attendees, List<Guid> roles) {
			List<string> emailsRequired = GetMeetingParticipantEmails(userConnection, activity.PrimaryColumnValue, roles);
			foreach (string emailItem in emailsRequired) {
				attendees.Add(emailItem);
			}
		}

		/// <summary>
		/// Validates exchange appointment before synchronization to bpmonline.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns><c>True</c> when appointment can be synchronized to bpmonline.
		/// Otherwise returns <c>false</c>.</returns>
		protected bool Validate(SyncContext context, LocalItem localItem) {
			return CheckNotDeleted(context, localItem) && CheckTypeAndICalUId(context) &&
				LockItem(context) && CheckOrganizer(context, localItem) &&
				CheckAppointmentContentDuplicates(context, localItem);
		}

		/// <summary>
		/// Fills activity instance with data from exchange appointment.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		protected void FillActivity(SyncContext context, LocalItem localItem) {
			_activitySyncEntity = localItem.Entities["Activity"][0];
			var exchangeAppointment = GetAppointment();
			Entity activity = _activitySyncEntity.Entity;
			var displayName = GetDisplayName();
			if (!CanModifyActivity(context, localItem)) {
				LogInfo(context, "FillActivity: current user can not modify Activity {0} (id {1})",
					displayName, activity.PrimaryColumnValue);
				return;
			}
			SetNewActivityProperties(_activitySyncEntity, context.UserConnection);
			LogInfo(context, "FillActivity SetNewActivityProperties ended for {0} appointment", displayName);
			SyncAllAttendees(context, activity, exchangeAppointment, localItem);
			LogInfo(context, "FillActivity SyncAllAttendees ended for {0} appointment", displayName);
			SetActivityProporties(context.UserConnection, activity, exchangeAppointment);
			LogInfo(context, "FillActivity SetActivityProporties ended for {0} appointment", displayName);
			var activityHash = ActivityUtils.GetActivityHash(activity, context.UserConnection.CurrentUser.TimeZone);
			UpdateExtraParameters(exchangeAppointment, activityHash);
			LogInfo(context, "FillActivity UpdateExtraParameters ended, ExtraParameters = {0}",
				_activitySyncEntity.ExtraParameters);
			FillReccuringMasterLocalItem(context, localItem);
			LogInfo(context, "FillActivity FillReccuringMasterLocalItem ended, ExtraParameters = {0}",
				_activitySyncEntity.ExtraParameters);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Verifies if the appointment was changed.
		/// </summary>
		/// <param name="activityHash">Activity hash.</param>
		/// <param name="syncEntity">The object of activity synchronization.</param>
		/// <returns>Flag indicating if the appointment was changed.</returns>
		public bool IsActivityChanged(string activityHash, SyncEntity syncEntity) {
			if (!((string)syncEntity.ExtraParameters).IsNullOrEmpty()) {
				JObject activityExtendProperty = JObject.Parse(syncEntity.ExtraParameters);
				var oldActivityHash = (string)activityExtendProperty["ActivityHash"];
				return activityHash != oldActivityHash;
			}
			return true;
		}

		/// <summary>
		/// Checks if exchange appointment was changed since last synchronization.
		/// </summary>
		/// <param name="appointment"><see cref="Exchange.Appointment"/> instance.</param>
		/// <param name="oldActivityHash">Last activity synchronization hash.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns><c>True</c> if exchange appointment was changed since last synchronization,
		/// <c>false</c> otherwise.</returns>
		public bool IsAppointmentChanged(Exchange.Appointment appointment, string oldActivityHash,
				UserConnection userConnection) {
			EntitySchema activitySchema = userConnection.EntitySchemaManager.GetInstanceByName("Activity");
			Entity activity = activitySchema.CreateEntity(userConnection);
			activity.SetColumnValue("StatusId", Guid.Empty);
			activity.SetColumnValue("Title", string.Empty);
			SetActivityProporties(userConnection, activity, appointment);
			TimeZoneInfo currentTimeZone = userConnection.CurrentUser.TimeZone;
			string activityHash = GetActivityHash(activity.GetTypedColumnValue<string>("Title"),
				activity.GetTypedColumnValue<string>("Location"), activity.GetTypedColumnValue<DateTime>("StartDate"),
				activity.GetTypedColumnValue<DateTime>("DueDate"), activity.GetTypedColumnValue<Guid>("PriorityId"),
				activity.GetTypedColumnValue<string>("Notes"), currentTimeZone);
			return (activityHash != oldActivityHash);
		}

		/// <summary>
		/// Verifies if the activity participant was changed.
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="lastSyncVersion">Last synchronization date.</param>
		/// <param name="remoteId">Unique identifier appointment in remote storage.</param>
		/// <param name="timeZoneInfo">Local time.</param>
		/// <returns>Flag indicating if the activity participantt was changed.</returns>
		public bool IsActivityParticipantChanged(UserConnection userConnection, DateTime lastSyncVersion,
			string remoteId, TimeZoneInfo timeZoneInfo) {
			DateTime lastSyncVersioUtc = TimeZoneInfo.ConvertTimeToUtc(lastSyncVersion, timeZoneInfo);
			Select selectActivityParticipant = new Select(userConnection)
				.Column(Func.Count("Id"))
				.From("SysSyncMetaData")
				.Where("RemoteId").IsEqual(Column.Parameter(remoteId))
				.And("SyncSchemaName").IsEqual(Column.Parameter("ActivityParticipant"))
				.And("Version").IsGreaterOrEqual(Column.Parameter(lastSyncVersioUtc)) as Select;
			return selectActivityParticipant.ExecuteScalar<int>() > 0;
		}

		/// <summary>Set appointment extra parameters in metadata.</summary>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		/// <param name="activityHash">Activity hash.</param>
		public void UpdateExtraParameters(Exchange.Appointment exchangeAppointment, string activityHash = null) {
			var userConnection = _activitySyncEntity.Entity.UserConnection;
			string remoteId = null;
			if (exchangeAppointment.Id != null) {
				remoteId = exchangeAppointment.Id.UniqueId;
			}
			if (userConnection != null && userConnection.GetIsFeatureEnabled("PrivateMeetings")) {
				UpdateExtraParameters(remoteId, activityHash,
					exchangeAppointment.Sensitivity == Exchange.Sensitivity.Private, exchangeAppointment.Subject);
			} else {
			UpdateExtraParameters(remoteId, activityHash);
		}
		}

		/// <summary>Set appointment extra parameters in metadata.</summary>
		/// <param name="exchangeAppointmentId">Exchange appointment item id.</param>
		/// <param name="activityHash">Activity hash.</param>
		/// <param name="isPrivate">Private meeting sign.</param>
		/// <param name="title">Appointment title.</param>
		public void UpdateExtraParameters(string exchangeAppointmentId = null, string activityHash = null,
			bool isPrivate = false, string title = "") {
			JObject activityExtendProperty;
			if (!((string)_activitySyncEntity.ExtraParameters).IsNullOrEmpty()) {
				activityExtendProperty = JObject.Parse(_activitySyncEntity.ExtraParameters);
			} else {
				activityExtendProperty = new JObject();
			}
			Activity activity = (Activity)_activitySyncEntity.Entity;
			if (exchangeAppointmentId != null) {
				activityExtendProperty["RemoteId"] = exchangeAppointmentId;
			}
			if (!string.IsNullOrEmpty(activityHash)) {
				activityExtendProperty["ActivityHash"] = activityHash;
			}
			activityExtendProperty["StatusId"] = activity.StatusId.ToString();
			activityExtendProperty["EndDate"] = activity.DueDate.ToString();
			var userConnection = activity.UserConnection;
			if (userConnection != null && userConnection.GetIsFeatureEnabled("PrivateMeetings")) {
				activityExtendProperty["IsPrivate"] = isPrivate;
				activityExtendProperty["Title"] = title;
			}
			ExchangeUtility.SetActivityExtraParameters(_activitySyncEntity, activityExtendProperty);
		}

		/// <summary>
		/// Calls <see cref="Exchange.Item.Load(Exchange.PropertySet)"/> method for <paramref name="exchangeAppointment"/>.
		/// </summary>
		/// <param name="exchangeAppointment"><see cref="Exchange.Appointment"/> instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		public virtual void LoadItemProperties(Exchange.Appointment exchangeAppointment) {
			exchangeAppointment.Load(_propertySet);
		}

		/// <summary>
		/// Sets appointment extend property. 
		/// </summary>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		/// <param name="activity">Activity item.</param>
		/// <param name="context">Synchronization context.</param>
		/// <returns>Status flag setting an extend property.
		/// </returns>
		public bool SetAppointmentExtendProperty(SyncContext context, Exchange.Appointment exchangeAppointment, Entity activity) {
			try {
				Object localId;
				bool isPropertyExists = exchangeAppointment.TryGetProperty(ExchangeUtilityImpl.LocalIdProperty, out localId);
				if (!isPropertyExists || Guid.Parse((string)localId) != activity.PrimaryColumnValue) {
					if (!isPropertyExists) {
						exchangeAppointment.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, activity.PrimaryColumnValue.ToString());
						LogInfo(context, "SetAppointmentExtendProperty set LocalIdProperty from activity {0}, {1} for appointment = {2}",
							activity.PrimaryColumnValue, activity.PrimaryDisplayColumnValue, RemoteId);
						exchangeAppointment.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite,
								Exchange.SendInvitationsOrCancellationsMode.SendToNone);
					}
					return true;
				}
				return false;
			} catch (Exception ex) {
				LogError(context, Action, SyncDirection.Upload,
					"[ExchangeAppointment.SetAppointmentExtendProperty]: An error occurred while setting the extend property in the Exchange: {0}",
					ex.Message);
				return false;
			}
		}

		/// <summary>
		/// Gets flag if the current contact has organizer Email. 
		/// </summary>
		/// <param name="userConnection">A instance of the current user connection.</param>
		/// <param name="currentContactId">Current contact Id.</param>
		/// <param name="organizerEmail">Organizers email of activity.</param>
		/// <returns>Flag if the current contact has organizer Email.
		/// </returns>
		public bool GetIsContactEmailExist(UserConnection userConnection, Guid currentContactId, string organizerEmail) {
			return ContactUtilities.GetContactEmails(userConnection, currentContactId)
				.Any(p => p.Equals(organizerEmail, StringComparison.OrdinalIgnoreCase)) ||
				GetCurrentUserMailboxEmails(userConnection).Any(p => p.Equals(organizerEmail, StringComparison.OrdinalIgnoreCase));
		}

		/// <summary>
		/// Sets properties for the <paramref name="activity"/>. 
		/// </summary>
		/// <param name="userConnection">An instance of the current user connection.</param>
		/// <param name="activity">Activity item.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		public virtual void SetActivityProporties(UserConnection userConnection, Entity activity,
				Exchange.Appointment exchangeAppointment) {
			if (userConnection.GetIsFeatureEnabled("PrivateMeetings") && exchangeAppointment.Sensitivity == Exchange.Sensitivity.Private) {
				SetActivityTitle(activity, new ExchangeUtilityImpl().GetPrivateAppointmentTitleLczValue(userConnection));
				activity.SetColumnValue("Location", string.Empty);
				activity.SetColumnValue("Notes", string.Empty);
			} else {
			string exchangeSubject = ActivityUtils.FixActivityTitle(exchangeAppointment.Subject, userConnection,
				exchangeMaxTitleLength);
			SetActivityTitle(activity, exchangeSubject);
			activity.SetColumnValue("Location", exchangeAppointment.Location);
				activity.SetColumnValue("Notes", exchangeAppointment.Body.Text);
			}
			activity.SetColumnValue("ShowInScheduler", true);
			DateTime startDate = exchangeAppointment.SafeGetValue<DateTime>(
				Exchange.AppointmentSchema.Start).GetUserDateTime(userConnection);
			activity.SetColumnValue("StartDate", startDate);
			activity.SetColumnValue("DueDate", exchangeAppointment.SafeGetValue<DateTime>(
				Exchange.AppointmentSchema.End).GetUserDateTime(userConnection));
			activity.SetColumnValue("ModifiedById", userConnection.CurrentUser.ContactId);
			activity.SetColumnValue("PriorityId", exchangeAppointment.Importance.GetActivityPriority());
			SetActivityStatus(activity, userConnection.CurrentUser.GetCurrentDateTime());
			activity.SetColumnValue("RemindToOwner", exchangeAppointment.SafeGetValue<bool>(
				Exchange.ItemSchema.IsReminderSet));
			if (activity.GetTypedColumnValue<bool>("RemindToOwner")) {
				activity.SetColumnValue("RemindToOwnerDate", startDate.AddMinutes(
					-exchangeAppointment.ReminderMinutesBeforeStart));
			} else {
				activity.SetColumnValue("RemindToOwnerDate", null);
			}
		}

		/// <summary>
		/// Sets properties for the <paramref name="exchangeAppointment"/>. 
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="activity">Activity instance.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		public void SetExchangeAppointmentProporties(SyncContext context, Entity activity,
				Exchange.Appointment exchangeAppointment) {
			UserConnection userConnection = context.UserConnection;
			if (Action == SyncAction.Update) {
				LoadItemProperties(exchangeAppointment);
			}
			if (Action == SyncAction.Create) {
				exchangeAppointment.Sensitivity = Exchange.Sensitivity.Normal;
			}
			if (!userConnection.GetIsFeatureEnabled("PrivateMeetings") || exchangeAppointment.Sensitivity != Exchange.Sensitivity.Private) {
			exchangeAppointment.Subject = activity.GetTypedColumnValue<string>("Title");
			exchangeAppointment.Location = activity.GetTypedColumnValue<string>("Location");
				exchangeAppointment.Body = new Exchange.MessageBody(Exchange.BodyType.HTML, activity.GetTypedColumnValue<string>("Notes"));
			}
			var startDate = activity.GetTypedColumnValue<DateTime>("StartDate");
			var timeZoneInfo = GetExchangeAppointmentTimeZoneInfo(userConnection);
			exchangeAppointment.Start = startDate;
			exchangeAppointment.StartTimeZone = timeZoneInfo;
			exchangeAppointment.End = activity.GetTypedColumnValue<DateTime>("DueDate");
			exchangeAppointment.EndTimeZone = timeZoneInfo;
			exchangeAppointment.Importance = (Exchange.Importance)ExchangeUtility
				.GetExchangeImportance(activity.GetTypedColumnValue<Guid>("PriorityId"));
			if (!userConnection.GetIsFeatureEnabled("SkipExchangeAppointmentReminder")) {
				exchangeAppointment.IsReminderSet = activity.GetTypedColumnValue<bool>("RemindToOwner");
			} else if (Action != SyncAction.Update) {
				exchangeAppointment.IsReminderSet = false;
			}
			var remindToOwnerDate = activity.GetTypedColumnValue<DateTime>("RemindToOwnerDate");
			if (remindToOwnerDate != DateTime.MinValue && remindToOwnerDate <= startDate) {
				TimeSpan duration = startDate - remindToOwnerDate;
				exchangeAppointment.ReminderMinutesBeforeStart = Convert.ToInt32(duration.TotalMinutes);
			} else {
				exchangeAppointment.ReminderMinutesBeforeStart = 0;
			}
			exchangeAppointment.LegacyFreeBusyStatus = Exchange.LegacyFreeBusyStatus.Busy;
			SetAppointmentParticipants(userConnection, activity, exchangeAppointment);
			exchangeAppointment.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, activity.PrimaryColumnValue.ToString());
		}

		/// <summary>Fills exchange appointment participants. Only exchange item organizer can modify attendee collection.</summary>
		/// <param name="userConnection">User connection.</param>
		/// <param name="activity">Activity instance.</param>
		/// <param name="exchangeAppointment">Exchange appointment item in external storage.</param>
		public void SetAppointmentParticipants(UserConnection userConnection, Entity activity, Exchange.Appointment exchangeAppointment) {
				_participantsRoles = _participantsRoles ?? ActivityUtils.GetParticipantsRoles(userConnection);
			ClearAttendees(userConnection, exchangeAppointment);
			AddAppointmentParticipants(userConnection, activity, exchangeAppointment.RequiredAttendees,
				new List<Guid> { _participantsRoles["Responsible"], _participantsRoles["Participant"] });
			AddAppointmentParticipants(userConnection, activity, exchangeAppointment.OptionalAttendees,
				new List<Guid> { _participantsRoles["OptionalParticipant"] });
		}

		/// <summary>
		/// Fills element synchronization in the local storage <paramref name="localItem"/> 
		/// the value of the element in the external storage.
		/// </summary>
		/// <param name="localItem">The element synchronization in the local storage.</param>
		/// <param name="context">Synchronization context.</param>
		public override void FillLocalItem(SyncContext context, ref LocalItem localItem) {
			LogInfo(context, "FillLocalItem started for {0} appointment", GetDisplayName());
			if (!Validate(context, localItem)) {
				LogInfo(context, "Validation failed for {0} appointment", GetDisplayName());
				return;
			}
			LogInfo(context, "Validation passed for {0} appointment", GetDisplayName());
			FillActivity(context, localItem);
			LogInfo(context, "FillLocalItem ended for {0} appointment", GetDisplayName());
		}

		/// <summary>
		/// Fills element in the remote storage from the element synchronization
		/// in the local storage.<paramref name="localItem"/>.
		/// </summary>
		/// <param name="localItem">The element synchronization in the local storage.</param>
		/// <param name="context">Synchronization context.</param>
		public override void FillRemoteItem(SyncContext context, LocalItem localItem) {
			LogInfo(context, "FillRemoteItem started");
			List<SyncEntity> syncEntities = localItem.Entities["Activity"];
			LogInfo(context, "FillRemoteItem get Activity from localItem.Entities done");
			_activitySyncEntity = syncEntities.FirstOrDefault();
			LogInfo(context, "FillRemoteItem set _activitySyncEntity done");
			if (_activitySyncEntity != null && _activitySyncEntity.State == SyncState.Deleted) {
				Action = SyncAction.Delete;
				return;
			}
			LogInfo(context, "FillRemoteItem _activitySyncEntity validate done");
			var activityTitle = _activitySyncEntity.Entity.GetTypedColumnValue<string>("Title");
			LogInfo(context, "FillRemoteItem started for {0} activity", activityTitle);
			var exchangeAppointment = (Exchange.Appointment)Item;
			if (Action == SyncAction.None || Action == SyncAction.Delete) {
				LogInfo(context, "Activity {0} sync action is {1}, fill remote item skipped.", activityTitle, Action);
				return;
			}
			Entity activity = GetEntityInstance<Entity>(context, localItem, "Activity");
			_activitySyncEntity = syncEntities.FirstOrDefault();
			activityTitle = activity.GetTypedColumnValue<string>("Title");
			if (_activitySyncEntity != null && _activitySyncEntity.Action == SyncAction.Create) {
				LogInfo(context, "Activity sync entity {0} sync action is {1}, fill remote item skipped.", activityTitle, _activitySyncEntity.Action);
				Action = SyncAction.None;
				return;
			}
			var currentTimeZone = context.UserConnection.CurrentUser.TimeZone;
			var activityHash = ActivityUtils.GetActivityHash(
				activity.GetTypedColumnValue<string>("Title"),
				activity.GetTypedColumnValue<string>("Location"),
				activity.GetTypedColumnValue<DateTime>("StartDate"),
				activity.GetTypedColumnValue<DateTime>("DueDate"),
				activity.GetTypedColumnValue<Guid>("PriorityId"),
				activity.GetTypedColumnValue<string>("Notes"),
				currentTimeZone);
			if ((!IsActivityChanged(activityHash, _activitySyncEntity) &&
				!IsActivityParticipantChanged(context.UserConnection, context.LastSyncVersion, RemoteId, currentTimeZone))
					|| IsOldActivity(activity.GetTypedColumnValue<DateTime>("DueDate"), context) ||
					GetEntityLockedForSync(activity.PrimaryColumnValue, context)) {
				LogInfo(context, "Activity {0} not changed, fill remote item skipped.", activityTitle, _activitySyncEntity.Action);
				Action = SyncAction.None;
				return;
			}
			SetExchangeAppointmentProporties(context, (Entity)activity, exchangeAppointment);
			LogInfo(context, "FillRemoteItem SetExchangeAppointmentProporties ended for {0} activity", activityTitle);
			SetSendNotification(context, activity, _activitySyncEntity);
			LogInfo(context, "FillRemoteItem SetSendNotification ended for {0} activity", activityTitle);
			UpdateExtraParameters(exchangeAppointment, activityHash);
			var activityId = _activitySyncEntity.Entity.PrimaryColumnValue;
			LogInfo(context, "Remote item properties: Subject = {0}, Location = {1}" +
				" Start = {2}, End = {3}, Importance = {4}, IsReminderSet = {5}, ReminderMinutesBeforeStart = {6}",
				exchangeAppointment.Subject,
				exchangeAppointment.Location,
				exchangeAppointment.Start,
				exchangeAppointment.End,
				exchangeAppointment.Importance,
				exchangeAppointment.IsReminderSet,
				exchangeAppointment.ReminderMinutesBeforeStart);
			LogInfo(context, "FillRemoteItem ended for activity title = {0}, id = {1}, remoteId = {2}",
				activityTitle, activityId, RemoteId);
		}

		/// <summary>
		/// Fills activity status. If due date in past, sets completed, else sets new status.
		/// </summary>
		/// <param name="activity">Activity instance.</param>
		/// <param name="DateTime">Current user date.</param>
		public void SetActivityStatus(Entity activity, DateTime currentUserDateTime) {
			Guid statusId = activity.GetIsColumnValueLoaded("StatusId") 
				? activity.GetTypedColumnValue<Guid>("StatusId") 
				: Guid.Empty;
			if (statusId.IsEmpty()) {
				DateTime dueDate = activity.GetTypedColumnValue<DateTime>("DueDate");
				if (dueDate > currentUserDateTime) {
					activity.SetColumnValue("StatusId", ActivityConsts.NewStatusUId);
				} else {
					activity.SetColumnValue("StatusId", ActivityConsts.CompletedStatusUId);
				}
			}
		}

		#endregion

	}

	#endregion

	#region Class: ExchangeAppointmentUtility

	/// <summary>
	/// It provides utility methods for working with the object synchronization
	/// appointment<see cref="ExchangeAppointment"/>.
	/// </summary>
	public static class ExchangeAppointmentUtility
	{

		#region Methods: Public

		/// <summary>
		/// Returns attendee email formated address, address format: <c>Name <Address></c>.
		/// </summary>
		/// <param name="attendeeEmailAddress">Attendee email address instance
		/// <see cref="Exchange.EmailAddress"/>.</param>
		/// <returns>Formated email address.</returns>
		public static string AttendeeToEmailAddressString(this Exchange.EmailAddress attendeeEmailAddress) {
			return attendeeEmailAddress == null ? string.Empty :
				string.Format("{0} <{1}>; ", attendeeEmailAddress.Name, attendeeEmailAddress.Address).TrimStart();
		}

		/// <summary>
		/// Combines emails from the collection <paramref name="attendeeCollection"/> to single line.
		/// </summary>
		/// <remarks>Address format <c>Name <Address>;</c></remarks>
		/// <param name="attendeeCollection"> Instance of the collection of participants.
		/// <see cref="Exchange.AttendeeCollection"/>.</param>
		/// <returns>A string of email addresses.</returns>
		public static string AttendeeToEmailAddressString(this Exchange.AttendeeCollection attendeeCollection) {
			var str = new StringBuilder();
			foreach (Exchange.EmailAddress attendeeEmailAddress in attendeeCollection) {
				str.Append(attendeeEmailAddress.AttendeeToEmailAddressString());
			}
			return str.ToString();
		}

		#endregion
	}

	#endregion
}