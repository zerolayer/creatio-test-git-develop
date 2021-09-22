namespace Terrasoft.Mail.Sender {

	using System;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeEmailClientFactory

	/// <summary> ############ ####### ######## ########.</summary>
	/// <remarks> ##### ######## ##### <see cref="EmailClientFactory"/>.</remarks>
	[DefaultBinding(typeof(EmailClientFactory))]
	public class ExchangeEmailClientFactory : EmailClientFactory {

		#region LocalizableStrings

		/// <summary>
		/// ############ ###### # ########## "## ############## ### #######.".
		/// </summary>
		public LocalizableString NotEmailTypeError {
			get {
				return new LocalizableString(GetResourceStorage(), "ExchangeEmailClientFactory",
					"LocalizableStrings.NotEmailTypeError.Value").ToString();
			}
		}

		#endregion

		#region Constructors: Public

		/// <summary>############## ##### ######### ###### <see cref="ExchangeEmailClientFactory"/>.</summary>
		/// <param name="userConnection">######### ################# ###########.</param>
		public ExchangeEmailClientFactory(UserConnection userConnection)
			: base(userConnection) { }

		#endregion

		#region Methods: Private

		private IResourceStorage GetResourceStorage() {
			return UserConnection.Workspace.ResourceStorage;
		}

		private EmailClientType GetEmailClientType(string senderEmailAddress, bool ignoreRights = false) {
			var helper = ClassFactory.Get<EmailMessageHelper>(
				new ConstructorArgument("userConnection", UserConnection));
			var mailboxESQ = helper.GetMailboxEsq(senderEmailAddress, ignoreRights);
			EntitySchemaQueryColumn typeColumn = mailboxESQ.AddColumn("MailServer.Type.Id");
			EntityCollection mailboxEntities = mailboxESQ.GetEntityCollection(UserConnection);
			if (mailboxEntities.Count == 0) {
				throw new EmailException("ErrorOnSend", new ExchangeUtilityImpl().GetMailboxDoesNotExistLczValue(UserConnection));
			}
			bool isExchange = ExchangeConsts.ExchangeMailServerTypeId == mailboxEntities[0].GetTypedColumnValue<Guid>(typeColumn.Name);
			return isExchange ? EmailClientType.Exchange : EmailClientType.Smtp;
		}

		private EmailClientType GetEmailClientType(Guid mailServerId) {
			var mailServerESQ = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "MailServer");
			EntitySchemaQueryColumn typeColumn = mailServerESQ.AddColumn("Type");
			IEntitySchemaQueryFilterItem idFilter = mailServerESQ
				.CreateFilterWithParameters(FilterComparisonType.Equal, "Id", mailServerId);
			mailServerESQ.Filters.Add(idFilter);
			EntityCollection mailboxEntities = mailServerESQ.GetEntityCollection(UserConnection);
			if (mailboxEntities.Count == 0) {
				throw new EmailException("ErrorOnSend",
					new ExchangeUtilityImpl().GetMailServerDoesNotExistLczValue(UserConnection));
			}
			bool isExchange = ExchangeConsts.ExchangeMailServerTypeId == mailboxEntities[0]
				.GetTypedColumnValue<Guid>(typeColumn.ValueQueryAlias);
			return isExchange ? EmailClientType.Exchange : EmailClientType.Smtp;
		}

		private EmailClientType GetEmailClientType(Credentials credentials) {
			return GetEmailClientType(credentials.ServerId);
		}

		#endregion

		#region Methdos: Public

		/// <summary>
		///########## ######### ######### ####### ######### ####.</summary>
		/// <param name="emailClientType">### ######### #######.</param>
		/// <returns>######### ######### <see cref="IEmailClient"/>.</returns>
		public IEmailClient CreateEmailClient(EmailClientType emailClientType, string senderEmailAddress) {
			IEmailClient emailClient = GetCertainEmailClient();
			if (emailClient != null) {
				return emailClient;
			}
			return GetEmailClient(emailClientType, senderEmailAddress);
		}

		private IEmailClient GetEmailClient(EmailClientType emailClientType, string senderEmailAddress) {
			switch (emailClientType) {
				case EmailClientType.Exchange:
					return new ExchangeClient(UserConnection);
				case EmailClientType.Smtp:
					return base.CreateEmailClient(senderEmailAddress);
				default:
					throw new EmailException("NotEmailType", NotEmailTypeError);
			}
		}

		/// <summary>
		/// ### ######### ######### ###### ########## ######### ######### #######.</summary>
		/// <param name="senderEmailAddress">######## ##### ###########.</param>
		/// <remarks>### ######### ####### ############ ## ########### ######.</remarks>
		/// <returns>######### ######### <see cref="IEmailClient"/>.</returns>
		public override IEmailClient CreateEmailClient(string senderEmailAddress) {
			EmailClientType emailClientType = GetEmailClientType(senderEmailAddress);
			return CreateEmailClient(emailClientType, senderEmailAddress);
		}

		/// <summary>
		/// Returns email client type for specified email.</summary>
		/// <param name="senderEmailAddress">Sender email.</param>
		/// <returns><see cref="IEmailClient"/> instance.</returns>
		/// <param name="ignoreRights">Flag that indicates whether to ignore rights.</param>
		public override IEmailClient CreateEmailClient(string senderEmailAddress, bool ignoreRights) {
			EmailClientType emailClientType = GetEmailClientType(senderEmailAddress, ignoreRights);
			return CreateEmailClient(emailClientType, senderEmailAddress);
		}

		/// <summary>
		/// Returns email client type for specified cridentials.</summary>
		/// <param name="credentials">Connection parameters.</param>
		/// <returns><see cref="IEmailClient"/> instance.</returns>
		public override IEmailClient CreateEmailClient(Credentials credentials) {
			IEmailClient emailClient = GetCertainEmailClient(credentials);
			if (emailClient != null) {
				return emailClient;
			}
			EmailClientType emailClientType = GetEmailClientType(credentials);
			switch (emailClientType) {
				case EmailClientType.Exchange:
					return new ExchangeClient(UserConnection, credentials);
				case EmailClientType.Smtp:
					return base.CreateEmailClient(credentials);
				default:
					throw new EmailException("NotEmailType", NotEmailTypeError);
			}

		}

		/// <summary>
		/// Returns an email client object that implements interface <see cref="IEmailClient"/>
		/// according to specified connection parameters.</summary>
		/// <param name="credentials"><see cref="EmailContract.DTO.Credentials"/> instance of connection credentials.</param>
		/// <returns>Instance of <see cref="IEmailClient"/>.</returns>
		public override IEmailClient CreateEmailClient(EmailContract.DTO.Credentials credentials) {
			IEmailClient emailClient = GetCertainEmailClient(credentials);
			if (emailClient != null) {
				return emailClient;
			}
			if (credentials.ServerTypeId == ExchangeConsts.ExchangeMailServerTypeId) {
				return new ExchangeClient(UserConnection, credentials);
			}
			if (credentials.ServerTypeId == ExchangeConsts.ImapMailServerTypeId) {
				return base.CreateEmailClient(credentials);
			} else {
				throw new EmailException("NotEmailType", NotEmailTypeError);
			}
		}

		#endregion
	}

	#endregion

	#region Enum: EmailClientType

	/// <summary>
	/// ### ######### #######.
	/// </summary>
	public enum EmailClientType {

		/// <summary>
		/// Exchange.
		/// </summary>
		Exchange = 1,

		/// <summary>
		/// Smtp.
		/// </summary>
		Smtp = 2
	}

	#endregion

}