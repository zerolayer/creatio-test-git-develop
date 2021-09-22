namespace Terrasoft.Configuration.ExchangeSyncService
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Net;
	using System.ServiceModel;
	using System.ServiceModel.Activation;
	using System.ServiceModel.Web;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Web.Common;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	/// <summary>
	/// Provides service methods for Exchange service interaction.
	/// </summary>
	[ServiceContract]
	[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
	public class ExchangeSyncService : BaseService
	{

		#region Methods: Private

		/// <summary>
		/// Fetch <see cref="MailboxSyncSettings"/> instance by <see cref="MailboxSyncSettings.SenderEmailAddress"/> column value,
		/// filtered using <paramref name="senderEmailAddress"/>. 
		/// <paramref name="settings"/> would be set with fetched instance.
		/// </summary>
		/// <param name="senderEmailAddress">Sender email address filter value.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="settings">Fetched <see cref="MailboxSyncSettings"/> instance.</param>
		/// <returns><c>True</c> if mailbox successfully fetched from database, <c>false</c> otherwise.</returns>
		private bool FetchSettingsEntity(string senderEmailAddress, UserConnection userConnection, out Entity settings) {
			EntitySchema schema = userConnection.EntitySchemaManager.GetInstanceByName("MailboxSyncSettings");
			Entity mailboxSyncSettings = schema.CreateEntity(userConnection);
			settings = mailboxSyncSettings;
			return mailboxSyncSettings.FetchFromDB("SenderEmailAddress", senderEmailAddress);
		}

		///<summary>Returns password for current user existing mailbox.</summary>
		/// <param name="senderEmailAddress">Email address.</param>
		/// <param name="userConnection">User connection instance.</param>
		/// <returns>Mail settings password for current user.</returns>
		private string GetExistingMailboxPassword(string senderEmailAddress, UserConnection userConnection) {
			Entity mailboxSyncSettings;
			if (FetchSettingsEntity(senderEmailAddress, userConnection, out mailboxSyncSettings)) {
				return mailboxSyncSettings.GetTypedColumnValue<string>("UserPassword");
			}
			return string.Empty;
		}

		///<summary>Returns can current mailbox use OAuth credentials.</summary>
		/// <param name="senderEmailAddress">Email address.</param>
		/// <param name="userConnection">User connection instance.</param>
		/// <returns><c>True</c> if mailbox can use  OAuth settings, <c>false</c> otherwise.</returns>
		private bool GetSettingsHasOauth(string senderEmailAddress, UserConnection userConnection) {
			Entity mailboxSyncSettings;
			if (FetchSettingsEntity(senderEmailAddress, userConnection, out mailboxSyncSettings)) {
				return mailboxSyncSettings.GetTypedColumnValue<Guid>("OAuthTokenStorageId").IsNotEmpty();
			}
			return false;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Sets exchange server sertificate validation callback.
		/// </summary>
		protected virtual void SetServerCertificateValidation() {
			ServicePointManager.ServerCertificateValidationCallback += new ExchangeUtilityImpl().ValidateRemoteCertificate;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Returns <see cref="ExchangeMailFolder"/> collection from Exchange server.
		/// </summary>
		/// <param name="mailServerId"><see cref="MailServer"/> unique id.</param>
		/// <param name="mailboxName">User mailbox name.</param>
		/// <param name="mailboxPassword">User mailbox password.</param>
		/// <param name="senderEmailAddress">User email address.</param>
		/// <param name="folderClassName">Exchange folders class name.</param>
		/// <returns><see cref="ExchangeMailFolder"/> collection.</returns>
		[OperationContract]
		[WebInvoke(UriTemplate = "GetMailboxFolders", ResponseFormat = WebMessageFormat.Json,
			BodyStyle = WebMessageBodyStyle.WrappedRequest)]
		public List<ExchangeMailFolder> GetMailboxFolders(string mailServerId, string mailboxName,
				string mailboxPassword, string senderEmailAddress, string folderClassName) {
			if (string.IsNullOrEmpty(mailboxPassword)) {
				mailboxPassword = GetExistingMailboxPassword(senderEmailAddress, UserConnection);
			}
			var credentials = new Mail.Credentials {
				UserName = mailboxName,
				UserPassword = mailboxPassword,
				ServerId = Guid.Parse(mailServerId),
				UseOAuth = GetSettingsHasOauth(senderEmailAddress, UserConnection)
			};
			SetServerCertificateValidation();
			var exchangeUtility = new ExchangeUtilityImpl();
			Exchange.ExchangeService service = exchangeUtility.CreateExchangeService(UserConnection, credentials,
				senderEmailAddress, true);
			var filterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.Or);
			var filter = new Exchange.SearchFilter.IsEqualTo(Exchange.FolderSchema.FolderClass, folderClassName);
			var nullfilter = new Exchange.SearchFilter.Exists(Exchange.FolderSchema.FolderClass);
			filterCollection.Add(filter);
			filterCollection.Add(new Exchange.SearchFilter.Not(nullfilter));
			string[] selectedFolders = null;
			var idPropertySet = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly);
			Exchange.Folder draftFolder = null;
			if (folderClassName == ExchangeConsts.NoteFolderClassName) {
				var inboxId = new Exchange.FolderId(Exchange.WellKnownFolderName.Inbox, senderEmailAddress);
				Exchange.Folder inboxFolder = Exchange.Folder.Bind(service, inboxId, idPropertySet);
				if (inboxFolder != null) {
					selectedFolders = new[] { inboxFolder.Id.UniqueId };
				}
				var draftsId = new Exchange.FolderId(Exchange.WellKnownFolderName.Drafts, senderEmailAddress);
				draftFolder = Exchange.Folder.Bind(service, draftsId, idPropertySet);
			}
			List<ExchangeMailFolder> folders = exchangeUtility.GetHierarhicalFolderList(service,
				Exchange.WellKnownFolderName.MsgFolderRoot, filterCollection, senderEmailAddress, selectedFolders);
			if (folders != null && folders.Any()) {
				folders[0] = new ExchangeMailFolder {
					Id = folders[0].Id,
					Name = senderEmailAddress,
					ParentId = string.Empty,
					Path = string.Empty,
					Selected = false
				};
				if (draftFolder != null) {
					folders.Remove(folders.FirstOrDefault(e => e.Path == draftFolder.Id.UniqueId));
				}
			}
			return folders;
		}

		#endregion
	}
}