namespace IntegrationV2
{
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeFoldersManger

	[DefaultBinding(typeof(IExchangeFoldersManger), Name = "FoldersManager")]
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
		/// Creates remote folders ids select for <paramref name="senderEmailAddress"/>.
		/// </summary>
		/// <param name="senderEmailAddress">Mailbox sender email address.</param>
		/// <returns><see cref="Select"/> instance.</returns>
		private Select GetFolderIdsSelect(string senderEmailAddress) {
			var select = new Select(UserConnection)
					.Column("MFC", "FolderPath")
				.From("MailboxFoldersCorrespondence").As("MFC")
					.InnerJoin("MailboxSyncSettings").As("MSS").On("MFC", "MailboxId").IsEqual("MSS", "Id")
				.Where("MSS", "SenderEmailAddress").IsEqual(Column.Parameter(senderEmailAddress)) as Select;
			return select;
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IExchangeFoldersManger.GetRemoteFolderIdsList(string)"/>
		public List<string> GetRemoteFolderIdsList(string senderEmailAddress) {
			var select = GetFolderIdsSelect(senderEmailAddress);
			var result = new List<string>();
			using (DBExecutor executor = UserConnection.EnsureDBConnection()){
				using (IDataReader dataReader = select.ExecuteReader(executor)) {
					while (dataReader.Read()) {
						result.Add(dataReader.GetColumnValue<string>("FolderPath"));
					}
				}
			}
			return result;
		}

		#endregion

	}

	#endregion

}
