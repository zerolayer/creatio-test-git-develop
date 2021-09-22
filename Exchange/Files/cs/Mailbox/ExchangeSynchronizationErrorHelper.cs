namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.ExchangeApi.Interfaces;
	using Terrasoft.Mail;

	#region Class: ExchangeSynchronizationErrorHelper

	/// <summary> Represents class for the process synchronization error.</summary>
	/// <remarks> Class substitute <see cref="SynchronizationErrorHelper"/> class.</remarks>
	[DefaultBinding(typeof(SynchronizationErrorHelper))]
	public class ExchangeSynchronizationErrorHelper : SynchronizationErrorHelper
	{

		#region Constructors: Public

		public ExchangeSynchronizationErrorHelper(UserConnection userConnection) : base(userConnection) {
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Stops synchronization process for the specific sender email.
		/// </summary>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="serverTypeId">Mail server type identifier.</param>
		protected override void StopSyncProcess(string senderEmailAddress, Guid serverTypeId) {
			if (serverTypeId == ExchangeConsts.ExchangeMailServerTypeId) {
				var exchangeUtility = ClassFactory.Get<IExchangeUtility>();
				exchangeUtility.RemoveAllSyncJob(UserConnection, senderEmailAddress, serverTypeId);
			} else {
				var parameters = new Dictionary<string, object> {
						{ "SenderEmailAddress", senderEmailAddress },
						{ "CurrentUserId", UserConnection.CurrentUser.Id }
					};
				var scheduler = ClassFactory.Get<IImapSyncJobScheduler>();
				scheduler.RemoveSyncJob(UserConnection, parameters);
			}
		}

		#endregion

	}

	#endregion

}