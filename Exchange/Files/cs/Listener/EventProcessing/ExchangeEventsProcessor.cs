namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using IntegrationApi.Email;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.Web.Http.Abstractions;

	#region Class: ExchangeEventsProcessor

	/// <summary>
	/// Class provides methods for exchange events processing.
	/// </summary>
	[DefaultBinding(typeof(IExchangeEventsProcessor), Name = "ExchangeEventsProcessor")]
	public class ExchangeEventsProcessor : BaseEmailEventsProcessor, IExchangeEventsProcessor
	{

		#region Constructors: Public

		public ExchangeEventsProcessor(AppConnection appConnection) : this(appConnection, HttpContext.HttpContextAccessor) {
		}

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeEventsProcessor" />.
		/// </summary>
		/// <param name="appConnection"><see cref="AppConnection"/> instance.</param>
		public ExchangeEventsProcessor(AppConnection appConnection, IHttpContextAccessor accessor): base(appConnection, accessor) {
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Starts emails synchronization.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Synchronization parameters.</param>
		protected virtual void StartSynchronization(UserConnection userConnection, IDictionary<string, object> parameters) {
			ValidateEvent(userConnection, parameters);
			Terrasoft.Core.Tasks.Task.StartNewWithUserConnection<ExchangeEmailEventExecutor, IDictionary<string, object>>(parameters);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Processes new exchange email event.
		/// Creates <see cref="ExchangeEmailEventExecutor"/> instance and executes with <paramref name="uniqueId"/>.
		/// </summary>
		/// <param name="emailEvent">New exchange email event.</param>
		public void ProcessNewEmail(ExchangeEmailEvent emailEvent) {
			var userName = emailEvent.SysAdminUnitName;
			Guid mailboxSyncSettingsId = new Guid(emailEvent.Id);
			var userConnection = CreateUserConnection(userName, emailEvent.TimeZoneId, mailboxSyncSettingsId);
			string senderEmailAddress = GetMailboxAddress(userConnection, mailboxSyncSettingsId);
			var parameters = new Dictionary<string, object> {
				{ "SenderEmailAddress", senderEmailAddress },
				{ "ItemIds", emailEvent.UniqueIds },
				{ "EventTimestamp", emailEvent.EventTimeStamp }
			};
			try {
				StartSynchronization(userConnection, parameters);
			} finally {
				userConnection?.Close(SessionEndMethod.Logout, false);
			}
		}

		#endregion

	}

	#endregion

}