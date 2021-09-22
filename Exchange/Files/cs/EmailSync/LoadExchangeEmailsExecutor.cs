namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Sync.Exchange;

	#region Class: LoadExchangeEmailsExecutor

	public class LoadExchangeEmailsExecutor : IJobExecutor
	{

		#region Properties: Public

		/// <summary>
		/// Result text message.
		/// </summary>
		public string ResultMessage {
			get;
			set;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Executes exchange emails synchronization.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Synchronization synchronization (user email address,
		/// create reminding sign etc.).</param>
		public virtual void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			string userEmailAddress = parameters["SenderEmailAddress"].ToString();
			if (string.IsNullOrEmpty(userEmailAddress)) {
				FormatResultMessage(new LocalizableString(userConnection.ResourceStorage, "LoadExchangeEmailsExecutor",
					"LocalizableStrings.NeedSetUserAddress.Value").ToString(), userEmailAddress);
				return;
			}
			string resultMessage;
			int localChangesCount, remoteChangesCount;
			ExchangeUtility.SyncExchangeItems(userConnection, userEmailAddress,
				() => new ExchangeEmailSyncProvider(userConnection, userEmailAddress),
				out resultMessage, out localChangesCount, out remoteChangesCount,
				ExchangeUtility.MailSyncProcessName);
		}

		/// <summary>
		/// Generates result text message.
		/// </summary>
		/// <param name="message">Result text message template.</param>
		/// <param name="userEmailAddress">User email address.</param>
		public virtual void FormatResultMessage(string message, string userEmailAddress) {
			string resultMessage = string.Format("{{\"key\": \"{0}\", \"message\": \"{1}\"}},",
				userEmailAddress, message);
			ResultMessage = string.Format("{{ \"Messages\": [{0}] }}", resultMessage);
		}

		#endregion

	}

	#endregion

}

