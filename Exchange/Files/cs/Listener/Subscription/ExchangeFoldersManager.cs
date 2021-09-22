namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.Sync.Exchange;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeFoldersManger

	[DefaultBinding(typeof(IExchangeFoldersManger), Name = "LegacyFoldersManager")]
	public class ExchangeFoldersManger : IExchangeFoldersManger
	{
		#region Fields: Protected

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		protected readonly UserConnection UserConnection;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Initializes a new instance of the <see cref="ExchangeFoldersManger"/> class.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		public ExchangeFoldersManger(UserConnection userConnection) {
			UserConnection = userConnection;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Binds exchange folders using <paramref name="folderRemoteIds"/>.
		/// </summary>
		/// <param name="service"><see cref="Microsoft.Exchange.WebServices.Data.ExchangeService"/> instance.</param>
		/// <param name="folderRemoteIds">Exchange folders unique identifiers collection.</param>
		/// <returns>Exchange folders collection.</returns>
		private List<Exchange.Folder> SafeBindFolders(Exchange.ExchangeService service, IEnumerable<string> folderRemoteIds) {
			List<Exchange.Folder> result = new List<Exchange.Folder>();
			foreach (string uniqueId in folderRemoteIds) {
				if (string.IsNullOrEmpty(uniqueId)) {
					continue;
				}
				try {
					Exchange.Folder folder = Exchange.Folder.Bind(service, new Exchange.FolderId(uniqueId));
					result.Add(folder);
				} catch (Exchange.ServiceResponseException) { }
			}
			return result;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="IExchangeFoldersManger.GetRemoteFolderIdsList"/>
		/// </summary>
		public List<string> GetRemoteFolderIdsList(string senderEmailAddress) {
			var service = new ExchangeUtilityImpl().CreateExchangeService(UserConnection, senderEmailAddress);
			var exchangeSettings = new EmailExchangeSettings(UserConnection, senderEmailAddress);
			List<Exchange.Folder> folders = new List<Exchange.Folder>();
			if (!exchangeSettings.LoadAll) {
				folders = SafeBindFolders(service, exchangeSettings.RemoteFolderUIds.Keys);
			}
			return folders.Count > 0 ? folders.Select(f => f.Id.UniqueId).ToList() : new List<string>();
		}

		#endregion
	}

	#endregion
}