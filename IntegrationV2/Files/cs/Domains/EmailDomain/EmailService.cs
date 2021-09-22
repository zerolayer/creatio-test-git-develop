namespace Terrasoft.EmailDomain
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.EmailDomain.Model;

	#region Class: EmailService

	/// <summary>
	/// Email model service implementation.
	/// </summary>
	[DefaultBinding(typeof(IEmailService))]
	public class EmailService : IEmailService
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="IEmailRepository"/> implementation instance.
		/// </summary>
		private readonly IEmailRepository _emailRepository;

		/// <summary>
		/// <see cref="IAttachmentRepository"/> implementation instance.
		/// </summary>
		private readonly IAttachmentRepository _attachmentRepository;

		/// <summary>
		/// <see cref="IActivityUtils"/> implementation instance.
		/// </summary>
		private readonly IActivityUtils _activityUtils;

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		#endregion

		#region Constructors: Public

		public EmailService(UserConnection uc) {
			_emailRepository = ClassFactory.Get<IEmailRepository>(new ConstructorArgument("uc", uc));
			_attachmentRepository = ClassFactory.Get<IAttachmentRepository>(new ConstructorArgument("uc", uc));
			_activityUtils = ClassFactory.Get<IActivityUtils>();
			_userConnection = uc;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates email model instance using <paramref name="emailDto"/>.
		/// </summary>
		/// <param name="emailDto"><see cref="Email"/> instance.</param>
		/// <returns>Email model instnace.</returns>
		private EmailModel CreateEmailModel(Email emailDto) {
			var utils = ClassFactory.Get<IActivityUtils>();
			var emailModel = new EmailModel() {
				From = emailDto.Sender,
				Subject = utils.FixActivityTitle(emailDto.Subject, _userConnection),
				SendDate = utils.GetSendDateFromTicks(_userConnection, emailDto.SendDateTimeStamp),
				IsHtmlBody = emailDto.IsHtmlBody,
				Headers = emailDto.Headers,
				To = emailDto.Recepients,
				Copy = emailDto.CopyRecepients,
				BlindCopy = emailDto.BlindCopyRecepients,
				Importance = emailDto.Importance,
				MessageId = emailDto.MessageId,
				InReplyTo = emailDto.InReplyTo,
				References = emailDto.References
			};
			emailModel.Attachments = CreateAttachments(emailDto, out string fixedBody);
			emailModel.Body = fixedBody;
			emailModel.OriginalBody = emailDto.Body;
			return emailModel;
		}

		/// <summary>
		/// Creates attachment model collection for <paramref name="emailDto"/>. Replaces inline attachments links in body.
		/// </summary>
		/// <param name="emailDto"><paramref name="emailDto"/> instnace.</param>
		/// <param name="fixedBody">Email body with replaced inline attachments links.</param>
		/// <returns>Attachment models collection.</returns>
		private List<AttachmentModel> CreateAttachments(Email emailDto, out string fixedBody) {
			var result = new List<AttachmentModel>();
			fixedBody = emailDto.Body;
			foreach (var attach in emailDto.Attachments) { 
				var attachModel = new AttachmentModel() {
					IsInline = attach.IsInline,
					Name = _activityUtils.GetAttachmentName(_userConnection, attach.Name),
					Id = Guid.NewGuid(),
					Data = attach.GetData()
				};
				if (attach.IsInline) {
					var url = _attachmentRepository.GetAttachmentLink(attachModel.Id);
					var cidUrl = string.Concat("cid:", attach.Id);
					if (fixedBody.Contains(cidUrl)) {
						fixedBody = fixedBody.Replace(cidUrl, url);
					} else {
						attachModel.IsInline = false;
					}
				}
				result.Add(attachModel);
			}
			return result;
		}

		/// <summary>
		/// Update attachment inline flag.
		/// </summary>
		/// <param name="email"><see cref="Email"/> instance.</param>
		private void UpdateAttacmentsInlineFlag(Email email) {
			foreach (var attachment in email.Attachments) {
				if (email.Body.Contains("cid:" + attachment.Id)) {
					_attachmentRepository.SetInline(Guid.Parse(attachment.Id));
				}
			}
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IEmailService.GetEmail(Guid))"/>
		public Email GetEmail(Guid activityId) {
			var email = _emailRepository.CreateEmail(activityId);
			var attachments = _attachmentRepository.GetAttachments(activityId);
			email.Attachments = attachments;
			return email;
		}

		/// <inheritdoc cref="IEmailService.SaveEmail(Email, Guid, string))"/>
		public void Save(Email email, Guid mailboxId = default(Guid), string syncSessionId = null) {
			var emailModel = CreateEmailModel(email);
			_emailRepository.Save(emailModel, mailboxId, syncSessionId);
		}

		/// <inheritdoc cref="IEmailService.Send(Email, Credentials)"/>
		public string Send(Email email, Credentials credentials) {
			var requestFactory = ClassFactory.Get<IHttpWebRequestFactory>();
			var emailProvider = ClassFactory.Get<IEmailProvider>(
				new ConstructorArgument("userConnection", _userConnection),
				new ConstructorArgument("requestFactory", requestFactory));
			UpdateAttacmentsInlineFlag(email);
			return emailProvider.Send(email, email.Attachments, credentials);
		}

		/// <inheritdoc cref="IEmailService.GetActivityIds(string)"/>
		public IEnumerable<Guid> GetActivityIds(string messageId) {
			return _emailRepository.GetHeaders(messageId).Select(em => em.Id);
		}

		#endregion

	}

	#endregion

}