namespace Terrasoft.EmailDomain
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using System.Security;
	using EmailContract;
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.EmailDomain.Model;
	using Terrasoft.IntegrationV2.Utils;
	using Terrasoft.Mail.Sender;

	#region Class: EmailRepository

	/// <summary>
	/// Email message model repository.
	/// </summary>
	[DefaultBinding(typeof(IEmailRepository))]
	internal class EmailRepository : IEmailRepository
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		/// <summary>
		/// <see cref="IActivityUtils"/> implementation instance.
		/// </summary>
		private readonly IActivityUtils _activityUtils = ClassFactory.Get<IActivityUtils>();

		/// <summary>
		/// <see cref="IAttachmentRepository"/> implementation instance.
		/// </summary>
		private readonly IAttachmentRepository _attachmentRepository;

		#endregion

		#region Constructors: Public

		public EmailRepository(UserConnection uc) {
			_userConnection = uc;
			_attachmentRepository = ClassFactory.Get<IAttachmentRepository>(new ConstructorArgument("uc", uc));
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates activity instance for <paramref name="email"/>.
		/// </summary>
		/// <param name="email">Email model instance.</param>
		/// <returns>Saved activity instance.</returns>
		private Entity CreateActivity(EmailModel email) {
			var activity = GetActivityEntity();
			activity.SetDefColumnValues();
			activity.SetColumnValue("Sender", email.From);
			activity.SetColumnValue("SendDate", email.SendDate);
			activity.SetColumnValue("Title", email.Subject);
			activity.SetColumnValue("TypeId", IntegrationConsts.EmailTypeId);
			activity.SetColumnValue("Body", email.Body);
			activity.SetColumnValue("IsHtmlBody", email.IsHtmlBody);
			activity.SetColumnValue("HeaderProperties", string.Join("\n", email.Headers));
			activity.SetColumnValue("OwnerId", _userConnection.CurrentUser.ContactId);
			activity.SetColumnValue("Recepient", string.Join(" ", email.To));
			activity.SetColumnValue("CopyRecepient", string.Join(" ", email.Copy));
			activity.SetColumnValue("BlindCopyRecepient", string.Join(" ", email.BlindCopy));
			activity.SetColumnValue("EmailSendStatusId", IntegrationConsts.EmailSentStatusId);
			activity.SetColumnValue("PriorityId", GetActivityPriority(email.Importance));
			activity.SetColumnValue("ActivityCategoryId", IntegrationConsts.EmailCategoryId);
			activity.SetColumnValue("StatusId", IntegrationConsts.ActivityCompletedStatusId);
			activity.SetColumnValue("DueDate", email.SendDate);
			activity.SetColumnValue("StartDate", email.SendDate);
			activity.SetColumnValue("MailHash", _activityUtils.GetEmailHash(_userConnection, email.SendDate, email.Subject,
				email.OriginalBody, _userConnection.CurrentUser.TimeZone));
			SaveActivity(activity, email);
			return activity;
		}

		/// <summary>
		/// Saves <paramref name="activity"/> to database.
		/// </summary>
		/// <param name="activity">Activity entity instance.</param>
		/// <param name="email">Email model instance.</param>
		private void SaveActivity(Entity activity, EmailModel email) {
			var emailId = GetActivityIdByMessageId(email.MessageId);
			if (emailId == Guid.Empty && ListenerUtils.GetIsFeatureDisabled(_userConnection, "SkipMailHashCheck")) {
				emailId = GetActivityIdByHash(email);
			}
			if (emailId != Guid.Empty) {
				activity.PrimaryColumnValue = emailId;
				email.Id = emailId;
			} else {
				activity.Save();
				email.Id = activity.PrimaryColumnValue;
				_attachmentRepository.SaveAttachments(email);
			}
		}

		/// <summary>
		/// Creates EmailMessageData instance for <paramref name="activity"/>.
		/// </summary>
		/// <param name="activity">Activity entity instance.</param>
		/// <param name="email">Email model instance.</param>
		/// <param name="mailboxId">Mailbox identifier.</param>
		/// <param name="syncSessionId">Synchronization session identifier.</param>
		private void CreateEmailMessageData(Entity activity, EmailModel email, Guid mailboxId, string syncSessionId) {
			var ticks = _activityUtils.GetSendDateTicks(_userConnection, activity);
			var userConnectionParam = new ConstructorArgument("userConnection", _userConnection);
			var helper = ClassFactory.Get<IEmailMessageHelper>(userConnectionParam);
			Dictionary<string, string> headers = new Dictionary<string, string>() {
				{ "MessageId", email.MessageId },
				{ "InReplyTo", email.InReplyTo },
				{ "SyncSessionId", syncSessionId },
				{ "References", email.References },
				{ "SendDateTicks", ticks.ToString() }
			};
			helper.CreateEmailMessage(activity, mailboxId, headers);
		}

		/// <summary>
		/// Returns <paramref name="importance"/> identifier.
		/// </summary>
		/// <param name="importance">Email message importance.</param>
		/// <returns>Message importance identifier.</returns>
		private Guid GetActivityPriority(EmailImportance importance) {
			return EmailPriorityConverter.GetActivityPriority((int)importance);
		}

		/// <summary>
		/// Returns activity identifier for <paramref name="messageId"/>.
		/// </summary>
		/// <param name="messageId">Email message identifier.</param>
		/// <returns>Activity identifier.</returns>
		private Guid GetActivityIdByMessageId(string messageId) {
			var headers = GetHeaders(messageId);
			return headers.Any() ? headers.First().Id : Guid.Empty;
		}

		/// <summary>
		/// Returns activity identifier using mail hash for <paramref name="email"/>.
		/// </summary>
		/// <param name="messageId"><see cref="EmailModel"/> instance.</param>
		/// <returns>Activity identifier.</returns>
		private Guid GetActivityIdByHash(EmailModel email) {
			List<Guid> emailIds = _activityUtils.GetExistingEmaisIds(_userConnection, email.SendDate, email.Subject,
				email.OriginalBody, _userConnection.CurrentUser.TimeZone);
			return emailIds.Any() ? emailIds.First() : Guid.Empty;
		}

		/// <summary>
		/// Creats email message headers select.
		/// </summary>
		/// <param name="messageId">Email message identifier.</param>
		/// <returns><see cref="Select"/> instance.</returns>
		private Select GetEmailHeaderSelect(string messageId) {
			return new Select(_userConnection)
				.Column("ActivityId")
				.Column("MessageId")
				.Column("InReplyTo")
				.Column("References")
			.From("EmailMessageData")
			.Where("MessageId").IsEqual(Column.Parameter(messageId)) as Select;
		}

		/// <summary>
		/// Create empty activity <see cref="Entity"/>.
		/// </summary>
		/// <returns>Activity<see cref="Entity"/>.</returns>
		private Entity GetActivityEntity() {
			var schema = _userConnection.EntitySchemaManager.GetInstanceByName("Activity");
			return schema.CreateEntity(_userConnection);
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IEmailRepository.CreateEmail(Guid)"/>
		public Email CreateEmail(Guid activityId) {
			var email = new Email {
				Id = activityId.ToString()
			};
			var activityEntity = GetActivityEntity();
			if (activityEntity.FetchFromDB(activityId)) {
				email.Body = activityEntity.GetTypedColumnValue<string>("Body");
				email.IsHtmlBody = activityEntity.GetTypedColumnValue<bool>("IsHtmlBody");
			} else {
				throw new SecurityException(string.Format(
					new LocalizableString("Terrasoft.Core", "EntitySchema.Exception.NoRightForRead"),
					activityEntity.SchemaName));
			}
			return email;
		}

		/// <inheritdoc cref="IEmailRepository.Save(EmailModel, Guid, string)"/>
		public void Save(EmailModel email, Guid mailboxId = default(Guid), string syncSessionId = null) {
			var activity = CreateActivity(email);
			CreateEmailMessageData(activity, email, mailboxId, syncSessionId);
		}

		/// <inheritdoc cref="IEmailRepository.GetHeaders(string)"/>
		public IEnumerable<EmailModelHeader> GetHeaders(string messageId) {
			var result = new List<EmailModelHeader>();
			if (messageId.IsNullOrEmpty()) {
				return result;
			}
			var select = GetEmailHeaderSelect(messageId);
			using (DBExecutor dbExecutor = _userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
					while (reader.Read()) {
						result.Add(new EmailModelHeader { 
							Id = reader.GetColumnValue<Guid>("ActivityId"),
							MessageId = reader.GetColumnValue<string>("MessageId"),
							InReplyTo = reader.GetColumnValue<string>("InReplyTo"),
							References = reader.GetColumnValue<string>("References")
						});
					}
				}
			}
			return result;
		}

		#endregion

	}

	#endregion

}
