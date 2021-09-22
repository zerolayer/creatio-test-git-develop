namespace Terrasoft.Sync.Exchange
{
	using System;
	using Terrasoft.Core;

	#region Class: ExchangeEmailEventsProvider

	public class ExchangeEmailEventsProvider : ExchangeEmailSyncProvider
	{

		#region Constructors: Public

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEmailEventsProvider" /> with passed synchronization settings.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="userSettings"><see cref="EmailExchangeSettings"/> instance.</param>
		public ExchangeEmailEventsProvider(UserConnection userConnection, string senderEmailAddress, EmailExchangeSettings userSettings = null)
			: base(userConnection, senderEmailAddress, userSettings) {

		}

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEmailEventsProvider" />.
		/// </summary>
		/// <param name="timeZoneInfo">Current timezone.</param>
		/// <param name="senderEmailAddress">The mailing address of the synchronization.</param>
		public ExchangeEmailEventsProvider(string senderEmailAddress, TimeZoneInfo timeZoneInfo)
			: base(senderEmailAddress, timeZoneInfo) {

		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="ExchangeEmailSyncProvider.LoadSyncItem(SyncItemSchema, string)"/>
		/// </summary>
		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, string id) {
			Service = Service ?? InitializeService(_userConnection);
			return base.LoadSyncItem(schema, id);
		}

		#endregion

	}

	#endregion

}