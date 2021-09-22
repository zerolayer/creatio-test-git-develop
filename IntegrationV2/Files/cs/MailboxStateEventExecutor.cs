namespace IntegrationV2
{
	using System;
	using System.Collections.Generic;
	using IntegrationApi.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.Common.Json;
	using Terrasoft.IntegrationV2.Logging.Interfaces;
	using Terrasoft.Common;

	#region Class: MailboxStateEventExecutor

	/// <summary>
	/// Mailbox state event handler action.
	/// </summary>
	public class MailboxStateEventExecutor : IJobExecutor
	{

		#region Fields: Private

		private ISynchronizationLogger _log;

		#endregion

		#region Methods: Private

		/// <summary>
		/// Gets ignore errors retry count flag.
		/// </summary>
		/// <param name="parameters">Action parameters collection.</param>
		/// <returns><c>True</c> if current error ignores retry counter. Otherwise returns <c>false</c>.</returns>
		private bool GetIgnoreRetryCount(IDictionary<string, object> parameters) {
			if (!parameters.ContainsKey("Context")) {
				return false;
			}
			var rawContext = parameters["Context"]?.ToString();
			if (rawContext.IsNullOrEmpty()) {
				return false;
			}
			try {
				var context = Json.Deserialize<Dictionary<string, object>>(rawContext);
				return context.ContainsKey("UserRequested") && (bool)context["UserRequested"];
			} catch (Exception e) {
				_log.Error($"Context deserialization failed", e);
				return false;
			}
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IJobExecutor.Execute(UserConnection userConnection, IDictionary{string, object})"/>
		public void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			_log = ClassFactory.Get<ISynchronizationLogger>(new ConstructorArgument("userId", userConnection.CurrentUser.Id));
			var senderEmailAddress = parameters["SenderEmailAddress"].ToString();
			_log.Info($"ConnectionEventExecutor.Execute for {senderEmailAddress} started");
			var helper = ClassFactory.Get<ISynchronizationErrorHelper>
				(new ConstructorArgument("userConnection", userConnection));
			var ignoreRetryCount = GetIgnoreRetryCount(parameters);
			if ((bool)parameters["Avaliable"]) {
				helper.CleanUpSynchronizationError(senderEmailAddress);
			} else {
				helper.ProcessSynchronizationError(senderEmailAddress,
					parameters["ExceptionClassName"].ToString(), parameters["ExceptionMessage"].ToString(), ignoreRetryCount);
			}
			_log.Info($"ConnectionEventExecutor.Execute for {senderEmailAddress} finished");
		}

		#endregion

	}

	#endregion

}
