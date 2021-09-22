namespace IntegrationV2
{
	using System;
	using System.Collections.Generic;
	using EmailContract.DTO;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.EmailDomain;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.Mail;
	using Terrasoft.Mail.Sender;
	using Credentials = EmailContract.DTO.Credentials;
	using Mail = Terrasoft.Mail;

	#region Class: EmailClient

	/// <summary>Represents email client class.</summary>
	[DefaultBinding(typeof(IEmailClient), Name = "EmailClient")]
	public class EmailClient : IEmailClient
	{

		#region Fields: Private

		protected readonly UserConnection _userConnection;
		private readonly Credentials _credentials;
		private readonly IEmailService _emailService;
		private Mailbox _mailbox;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// <see cref="EmailClient"/> ctor.
		/// </summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		public EmailClient(UserConnection userConnection) {
			_userConnection = userConnection;
			_emailService = ClassFactory.Get<IEmailService>(
				new ConstructorArgument("uc", _userConnection));
		}

		/// <summary>
		/// <see cref="EmailClient"/> ctor.
		/// </summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		/// <param name="credentials">Short email connection credentials.</param>
		public EmailClient(UserConnection userConnection, Mail.Credentials credentials)
			: this(userConnection) {
			var mailServer = GetMailServer(credentials.ServerId);
			_credentials = new Credentials {
				UserName = credentials.UserName,
				Password = credentials.UserPassword,
				UseOAuth = credentials.UseOAuth,
				ServiceUrl = mailServer.OutgoingServerAddress,
				ServerTypeId = mailServer.TypeId,
				Port = mailServer.OutgoingPort,
				UseSsl = mailServer.OutgoingUseSsl
			};
		}

		/// <summary>
		/// <see cref="EmailClient"/> ctor.
		/// </summary>
		/// <param name="userConnection">An instance of the user connection.</param>
		/// <param name="emailCredentials">Full email connection credentials.</param>
		public EmailClient(UserConnection userConnection, Credentials emailCredentials)
			: this(userConnection) {
			_credentials = emailCredentials;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Set email headers.
		/// </summary>
		/// <param name="email"><see cref="Email"/> instance.</param>
		/// <param name="headerProperties">List of <see cref="EmailMessageHeader"/>.</param>
		private void SetEmailHeaders(Email email, List<EmailMessageHeader> headerProperties) {
			var headers = new List<string>();
			if (headerProperties != null) {
				foreach (var property in headerProperties) {
					headers.Add(string.Concat(property.Name, "=", property.Value));
				}
			}
			email.Headers = headers;
		}

		/// <summary>
		/// Fills recipients from <see cref="EmailMessage"/> to <see cref="Email"/>.
		/// </summary>
		/// <param name="email"><see cref="Email"/> instance.</param>
		/// <param name="emailMessage"><see cref="EmailMessage"/> instance.</param>
		private void SetEmailRecipients(Email email, EmailMessage emailMessage) {
			FillMessageRecipientsCollection(email.Recepients, emailMessage.To);
			FillMessageRecipientsCollection(email.CopyRecepients, emailMessage.Cc);
			FillMessageRecipientsCollection(email.BlindCopyRecepients, emailMessage.Bcc);
		}

		/// <summary>
		/// Fills <paramref name="collection"/> recipients collection with <paramref name="values"/>.
		/// </summary>
		/// <param name="collection">Exchange email recipients collection.</param>
		/// <param name="values">Recipients values.</param>
		private void FillMessageRecipientsCollection(List<string> collection, List<string> values) {
			foreach (var address in values) {
				collection.Add(ExtractEmailAddress(address));
			}
		}

		/// <summary>
		/// Tries to find <paramref name="address"/> email address in string.
		/// </summary>
		/// <param name="address">Email address containing string.</param>
		/// <returns>Email address.</returns>
		private string ExtractEmailAddress(string address) {
			int first = address.IndexOf('<');
			if (first == -1) {
				return address.Trim();
			}
			first += 1;
			int last = address.LastIndexOf('>');
			int count = last - first;
			address = (count >= 0) ? address.Substring(first, count) : address.Substring(first);
			return address.Trim();
		}

		/// <summary>
		/// Returns <see cref="Email"/> instance.</summary>
		/// <param name="emailMessage"><see cref="EmailMessage"/> instance.</param>
		/// <param name="ignoreRights">Flag that indicates whether to ignore rights.</param> 
		/// <returns><see cref="Email"/> instance.</returns>
		private Email GetEmail(EmailMessage emailMessage, bool ignoreRights = false) {
			var email = new Email {
				Id = emailMessage.GetMessageId(),
				Subject = emailMessage.Subject,
				Body = emailMessage.Body,
				Sender = GetSender(emailMessage, ignoreRights),
				Importance = (EmailContract.EmailImportance)emailMessage.Priority,
				IsHtmlBody = emailMessage.IsHtmlBody
			};
			SetEmailRecipients(email, emailMessage);
			SetEmailHeaders(email, emailMessage.HeaderProperties);
			SetAttachments(email, emailMessage.Attachments);
			return email;
		}

		/// <summary>
		/// Returns email sender value.</summary>
		/// <param name="emailMessage"><see cref="EmailMessage"/> instance.</param>
		/// <param name="ignoreRights">Flag that indicates whether to ignore rights.</param> 
		/// <returns>Email sender value.</returns>
		private string GetSender(EmailMessage emailMessage, bool ignoreRights = false) {
			if (_credentials != null) {
				return emailMessage.From;
			}
			var mailbox = GetMailbox(emailMessage.From, ignoreRights);
			return mailbox.GetSender();
		}

		/// <summary>
		/// Fill <see cref="Email.Attachments"/> collection.</summary>
		/// <param name="emailMessage"><see cref="Email"/> instance.</param>
		/// <param name="emailMessage"><see cref="EmailAttachment"/> collection.</param>
		private void SetAttachments(Email email, List<EmailAttachment> attachments) {
			var attachmentsDto = new List<Attachment>();
			foreach (EmailAttachment attachment in attachments) {
				var attachmentDto = new Attachment {
					Id = attachment.Id.ToString(),
					Name = attachment.Name,
					IsInline = attachment.IsContent
				};
				attachmentDto.SetData(attachment.Data);
				attachmentsDto.Add(attachmentDto);
			}
			email.Attachments = attachmentsDto;
		}

		/// <summary>
		/// Get send email credentials.
		/// </summary>
		/// <param name="from">Sender email address.</param>
		/// <param name="ignoreRights">Ignore mailbox rights flag.</param>
		/// <param name="useForSynchronization">Sign is synchronization mode or not.</param>
		/// <returns><see cref="Credentials"/> instance.</returns>
		private Credentials GetCredentials(string from, bool ignoreRights, bool useForSynchronization = true) {
			return _credentials == null
				? GetEmailCredentialsFromMailbox(from, ignoreRights, useForSynchronization)
				: GetEmailCredentials(from);
		}

		/// <summary>
		/// Returns <see cref="Credentials"/> instance.</summary>
		/// <param name="from">Sender email.</param>
		/// <param name="ignoreRights">Ignore mailbox rights flag.</param>
		/// <param name="useForSynchronization">Sign is synchronization mode or not.</param>
		/// <returns><see cref="Credentials"/> instance.</returns>
		private Credentials GetEmailCredentialsFromMailbox(string from, bool ignoreRights,
				bool useForSynchronization = true) {
			var mailbox = GetMailbox(from, ignoreRights);
			return mailbox.ConvertToSynchronizationCredentials(_userConnection, useForSynchronization);
		}

		/// <summary>
		/// Returns <see cref="Mailbox"/> instance.</summary>
		/// <param name="from">Sender email.</param>
		/// <param name="ignoreRights">Ignore mailbox rights flag.</param>
		/// <returns><see cref="Mailbox"/> instance.</returns>
		private Mailbox GetMailbox(string from, bool ignoreRights) {
			if (_mailbox != null) {
				return _mailbox;
			}
			var mailboxService = ClassFactory.Get<IMailboxService>(new ConstructorArgument("uc", _userConnection));
			_mailbox = mailboxService.GetMailboxBySenderEmailAddress(from, !ignoreRights, false);
			return _mailbox;
		}

		/// <summary>
		/// Returns <see cref="Credentials"/> instance from <see cref="Mail.Credentials"/> instance.</summary>
		/// <param name="from">Sender email.</param>
		/// <returns><see cref="Credentials"/> instance.</returns>
		private Credentials GetEmailCredentials(string from) {
			var credentials = _credentials;
			credentials.SenderEmailAddress = from;
			return credentials;
		}

		/// <summary>
		/// Returns mail server instance.</summary>
		/// <param name="serverId">Mail server identifier.</param>
		/// <returns>Mail server instance.</returns>
		private MailServer GetMailServer(Guid serverId) {
			var service = ClassFactory.Get<IMailServerService>(new ConstructorArgument("uc", _userConnection));
			return service.GetServer(serverId, false);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Preparing email by <paramref name="emailMessage"/> data and send it.</summary>
		/// <param name="emailMessage">Email message data.</param>
		/// <param name="ignoreRights">Flag that indicates whether to ignore rights.</param>
		public void Send(EmailMessage emailMessage, bool ignoreRights = false) {
			var emailDto = GetEmail(emailMessage, ignoreRights);
			var credentials = GetCredentials(emailMessage.From, ignoreRights, false);
			var sendResult = _emailService.Send(emailDto, credentials);
			if (sendResult.IsNotNullOrEmpty()) {
				throw new EmailException("ErrorOnSend", sendResult);
			}
		}

		#endregion

	}

	#endregion

}
