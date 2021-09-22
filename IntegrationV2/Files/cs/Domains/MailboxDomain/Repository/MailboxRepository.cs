namespace IntegrationV2.MailboxDomain.Repository
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using IntegrationApi.MailboxDomain.Model;
	using IntegrationV2.MailboxDomain.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Utils;

	#region Class: MailboxRepository

	/// <summary>
	/// Mailbox repository implementation.
	/// </summary>
	[DefaultBinding(typeof(IMailboxRepository))]
	internal class MailboxRepository : BaseRepository, IMailboxRepository
	{

		#region Fields: Private

		private IMailServerRepository _mailServerRepository;

		private IMailboxFolderRepository _mailboxFolderRepository;

		#endregion

		#region Constructors: Public

		public MailboxRepository(UserConnection uc) {
			UserConnection = uc;
			CacheName = "MailboxList";
			_mailServerRepository = ClassFactory.Get<IMailServerRepository>(new ConstructorArgument("uc", uc));
			_mailboxFolderRepository = ClassFactory.Get<IMailboxFolderRepository>(new ConstructorArgument("uc", uc));
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates mailbox data query.
		/// </summary>
		/// <param name="userMailboxesOnly">Select only user mailboxes flag.</param>
		/// <returns><see cref="EntitySchemaQuery"/> instance.</returns>
		private EntitySchemaQuery GetMailboxesQuery(bool userMailboxesOnly) {
			var esq = new EntitySchemaQuery(UserConnection.EntitySchemaManager, "MailboxSyncSettings");
			esq.UseAdminRights = userMailboxesOnly;
			esq.PrimaryQueryColumn.IsAlwaysSelect = true;
			esq.AddColumn("SenderEmailAddress");
			esq.AddColumn("SenderDisplayValue");
			esq.AddColumn("UserName");
			esq.AddColumn("UserPassword");
			esq.AddColumn("OAuthTokenStorage");
			esq.AddColumn("SysAdminUnit");
			esq.AddColumn("CreatedBy");
			esq.AddColumn("IsShared");
			esq.AddColumn("MailServer");
			esq.AddColumn("EnableMailSynhronization");
			esq.AddColumn("SendEmailsViaThisAccount");
			esq.AddColumn("SynchronizationStopped");
			esq.AddColumn("PersonalMetrics");
			esq.AddColumn("MailSyncPeriod");
			esq.AddColumn("AutomaticallyAddNewEmails");
			esq.AddColumn("ExchangeAutoSyncActivity");
			esq.AddColumn("ExchangeAutoSynchronization");
			esq.AddColumn("EmailsCyclicallyAddingInterval");
			esq.AddColumn("LoadAllEmailsFromMailBox");
			esq.AddColumn("ErrorCode");
			return esq;
		}

		/// <summary>
		/// Creates mailbox model instance.
		/// </summary>
		/// <param name="entity">Mailbox entity.</param>
		/// <param name="mailServers">Mail servers models collection.</param>
		/// <param name="folders">Mail folders collection.</param>
		/// <returns><see cref="Mailbox"/> instance.</returns>
		private Mailbox CreateMailbox(Entity entity, IEnumerable<MailServer> mailServers, IEnumerable<MailboxFolder> folders) {
			var mailServerId = entity.GetTypedColumnValue<Guid>("MailServerId");
			var mailboxFolders = folders.Where(f => f.MailboxId.Equals(entity.PrimaryColumnValue));
			var mailServer = mailServers.First(ms => ms.Id.Equals(mailServerId));
			return CreateMailbox(entity, mailServer, mailboxFolders);
		}

		/// <summary>
		/// Creates mailbox model instance.
		/// </summary>
		/// <param name="entity">Mailbox entity.</param>
		/// <param name="mailServer">Mail server entity.</param>
		/// <param name="mailboxFolders">Mailbox folders collection.</param>
		/// <returns><see cref="Mailbox"/> instance.</returns>
		private Mailbox CreateMailbox(Entity entity, MailServer mailServer, IEnumerable<MailboxFolder> mailboxFolders) {
			var mailbox = new Mailbox(entity, mailServer);
			mailbox.AddFolders(mailboxFolders);
			return mailbox;
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IMailboxRepository.GetAll(bool)"/>
		public IEnumerable<Mailbox> GetAll(bool userMailboxesOnly = true, bool useForSynchronization = true) {
			object store = GetCache();
			IEnumerable<Mailbox> mailboxList = null;
			if (store != null) {
				mailboxList = store as IEnumerable<Mailbox>;
			}
			var mailServers = _mailServerRepository.GetAll(useForSynchronization);
			var folders = _mailboxFolderRepository.GetAll();
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && mailboxList != null) {
				return mailboxList;
			} else {
				var mailboxesQuery = GetMailboxesQuery(userMailboxesOnly);
				var mailboxes = new List<Mailbox>();
				foreach (var mailbox in mailboxesQuery.GetEntityCollection(UserConnection)) {
					mailboxes.Add(CreateMailbox(mailbox, mailServers, folders));
				}
				if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && !userMailboxesOnly) {
					SetCache(mailboxes);
				}
				return mailboxes;
			}
		}

		/// <inheritdoc cref="IMailboxRepository.GetById(Guid)"/>
		public Mailbox GetById(Guid mailboxId, bool userMailboxesOnly = true) {
			var mailboxList = GetCache() as List<Mailbox>;
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && mailboxList != null
					&& mailboxList.Any(x => x.Id == mailboxId)) {
				return mailboxList.FirstOrDefault(x => x.Id == mailboxId);
			} else {
				var mailboxesQuery = GetMailboxesQuery(userMailboxesOnly);
				var mailboxEntity = mailboxesQuery.GetEntity(UserConnection, mailboxId);
				if (mailboxEntity == null) {
					return null;
				}
				var mailServerId = mailboxEntity.GetTypedColumnValue<Guid>("MailServerId");
				var mailServer = _mailServerRepository.GetById(mailServerId);
				var mailboxFolders = _mailboxFolderRepository.GetByMailboxId(mailboxId);
				var mailbox = CreateMailbox(mailboxEntity, mailServer, mailboxFolders);
				return mailbox;
			}
		}

		#endregion

	}

	#endregion

}
