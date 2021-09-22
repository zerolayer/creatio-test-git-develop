namespace IntegrationV2.MailboxDomain.Repository
{
	using System;
	using System.Collections.Generic;
	using IntegrationApi.MailboxDomain.Model;
	using IntegrationV2.MailboxDomain.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Factories;
	using System.Linq;
	using Terrasoft.IntegrationV2.Utils;

	#region Class: MailServerRepository

	/// <summary>
	/// Mail server repository implementation.
	/// </summary>
	[DefaultBinding(typeof(IMailServerRepository))]
	internal class MailServerRepository : BaseRepository, IMailServerRepository
	{

		#region Conctructors: Public

		public MailServerRepository(UserConnection uc) {
			UserConnection = uc;
			CacheName = "MailServerList";
		}

		#endregion

		#region Methods: Private

		private Select GetMailServersQuery() {
			return new Select(UserConnection)
					.Column("MS", "Id")
					.Column("MS", "IsExchengeAutodiscover")
					.Column("MS", "ExchangeEmailAddress")
					.Column("MS", "TypeId")
					.Column("MS", "AllowEmailDownloading")
					.Column("MS", "AllowEmailSending")
					.Column("MS", "Address")
					.Column("MS", "Port")
					.Column("MS", "SMTPServerAddress")
					.Column("MS", "SMTPPort")
					.Column("MS", "UseSSL")
					.Column("MS", "UseSSLforSending")
					.Column("MS", "Strategy")
					.Column("OAA", "ClientClassName")
					.Column("MS", "IsStartTls")
					.Column("MS", "SubscriptionTtl")
				.From("MailServer").As("MS")
				.LeftOuterJoin("OAuthApplications").As("OAA").On("MS", "OAuthApplicationId").IsEqual("OAA", "Id") as Select;
		}

		private void AddPrimaryFilter(Select query, Guid mailServerId) {
			query.Where("MS", "Id").IsEqual(Column.Parameter(mailServerId));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IMailServerRepository.GetAll"/>
		public IEnumerable<MailServer> GetAll(bool useForSynchronization = true) {
			object store = GetCache();
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && store != null) {
				return store as IEnumerable<MailServer>;
			} else {
				var query = GetMailServersQuery();
				var mailServers = new List<MailServer>();
				using (var dbExecutor = UserConnection.EnsureDBConnection()) {
					using (var reader = query.ExecuteReader(dbExecutor)) {
						while (reader.Read()) {
							mailServers.Add(new MailServer(reader, useForSynchronization));
						}
					}
				}
				if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached")) {
					SetCache(mailServers);
				}
				return mailServers;
			}
		}

		/// <inheritdoc cref="IMailServerRepository.GetById(Guid)"/>
		public MailServer GetById(Guid mailServerId) {
			MailServer mailserver = null;
			var mailServerList = GetCache() as List<MailServer>;
			if (ListenerUtils.GetIsFeatureEnabled(UserConnection, "IsMailboxSyncSettingsCached") && mailServerList != null
					&& mailServerList.Any(x => x.Id == mailServerId)) {
				mailserver = mailServerList.FirstOrDefault(x => x.Id == mailServerId);
			} else {
				var query = GetMailServersQuery();
				AddPrimaryFilter(query, mailServerId);
				using (var dbExecutor = UserConnection.EnsureDBConnection()) {
					using (var reader = query.ExecuteReader(dbExecutor)) {
						if (reader.Read()) {
							mailserver = new MailServer(reader);
						}
					}
				}
			}
			return mailserver;
		}

		#endregion

	}

	#endregion

}
