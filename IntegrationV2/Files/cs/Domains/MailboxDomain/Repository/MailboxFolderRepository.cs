namespace IntegrationV2.MailboxDomain.Repository
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using IntegrationApi.MailboxDomain.Model;
	using IntegrationV2.MailboxDomain.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Utils;

	#region Class: MailboxFolderRepository

	/// <summary>
	/// Mailbox folders storage repository implementation.
	/// </summary>
	[DefaultBinding(typeof(IMailboxFolderRepository))]
	internal class MailboxFolderRepository : BaseRepository, IMailboxFolderRepository
	{

		#region Constructors: Public

		public MailboxFolderRepository(UserConnection uc) {
			UserConnection = uc;
			CacheName = "MailboxFolderList";
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates mail folders query.
		/// </summary>
		/// <returns><see cref="Select"/> instance.</returns>
		private Select GetFoldersQuery() {
			return new Select(UserConnection)
					.Column("Id")
					.Column("FolderPath")
					.Column("MailboxId")
				.From("MailboxFoldersCorrespondence") as Select;
		}

		/// <summary>
		/// Create filter by mailbox.
		/// </summary>
		/// <param name="query"><see cref="Select"/> instance.</param>
		/// <param name="mailboxId">Mailbox identifier.</param>
		private void AddMailboxFilter(Select query, Guid mailboxId) {
			query.Where("MailboxId").IsEqual(Column.Parameter(mailboxId));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IMailboxFolderRepository.GetAll"/>
		public IEnumerable<MailboxFolder> GetAll() {
			object store = GetCache();
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && store != null) {
				return store as IEnumerable<MailboxFolder>;
			} else {
				var query = GetFoldersQuery();
				var mailboxFolders = new List<MailboxFolder>();
				using (var dbExecutor = UserConnection.EnsureDBConnection()) {
					using (var reader = query.ExecuteReader(dbExecutor)) {
						while (reader.Read()) {
							mailboxFolders.Add(new MailboxFolder(reader));
						}
					}
				}
				if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached")) {
					SetCache(mailboxFolders);
				}
				return mailboxFolders;
			}
		}

		/// <inheritdoc cref="IMailboxFolderRepository.GetByMailboxId(Guid)"/>
		public IEnumerable<MailboxFolder> GetByMailboxId(Guid mailboxId) {
			var mailFolderList = GetCache() as List<MailboxFolder>;
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && mailFolderList != null
					&& mailFolderList.Any(x => x.MailboxId == mailboxId)) {
				return mailFolderList.Where(x => x.MailboxId == mailboxId);
			} else {
				var query = GetFoldersQuery();
				AddMailboxFilter(query, mailboxId);
				var mailboxFolders = new List<MailboxFolder>();
				using (var dbExecutor = UserConnection.EnsureDBConnection()) {
					using (var reader = query.ExecuteReader(dbExecutor)) {
						while (reader.Read()) {
							mailboxFolders.Add(new MailboxFolder(reader));
						}
					}
				}
				return mailboxFolders;
			}
		}

		#endregion

	}

	#endregion

}
