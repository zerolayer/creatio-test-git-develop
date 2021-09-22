namespace Terrasoft.IntegrationV2.Utils
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using IntegrationApi.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Configuration;
	using Terrasoft.Core.Factories;
	using TS = Terrasoft.Web.Http.Abstractions;


	#region Class: ListenerUtils

	[DefaultBinding(typeof(ListenerUtils))]
	internal class ListenerUtils
	{

		#region Fields: Private

		private readonly UserConnection _uc;

		private readonly TS.IHttpContextAccessor _context;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// .ctor.
		/// </summary>
		/// <param name="uc"><see cref="UserConnection"/> instance.</param>
		/// <param name="context"><see cref="TS.IHttpContextAccessor"/> instance.</param>
		public ListenerUtils(UserConnection uc, TS.IHttpContextAccessor context) {
			_uc = uc;
			_context = context ?? TS.HttpContext.HttpContextAccessor;
		}

		#endregion

		#region Properties: Internal

		internal static int FailoverJobFilterMinutesOffset => -5;

		#endregion

		#region Methods: Private

		/// <summary>
		/// Create new value of BpmonlineExchangeEventsEndpointUrl system setting.
		/// </summary>
		/// <returns>New value of BpmonlineExchangeEventsEndpointUrl system setting.</returns>
		private string GetNewBpmonlineExchangeEventsEndpointUrl() {
			var request = _context?.GetInstance()?.Request;
			var urlPath = "ServiceModel/ExchangeListenerService.svc/ProcessFullEmail";
			return request != null ? GetUrl(request, urlPath) : string.Empty;
		}

		/// <summary>
		/// Combined url from <paramref name="request"/> and <paramref name="path"/>. 
		/// </summary>
		/// <param name="request"><see cref="HttpRequest"/> instance.</param>
		/// <param name="path">Url path.</param>
		/// <returns>Url.</returns>
		private string GetUrl(TS.HttpRequest request, string path) {
			var combinedPath = string.Concat(request.ApplicationPath?.TrimEnd('/'), "/", path);
			var port = request.Host.Port ?? -1;
			var uriBuilder = new UriBuilder(request.Scheme, request.Host.Host, port, combinedPath);
			return uriBuilder.Uri.ToString();
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Returns is feature enabled for <paramref name="uc"/>.
		/// </summary>
		/// <param name="uc"><see cref="UserConnection"/> instance.</param>
		/// <param name="code">Feature code.</param>
		/// <returns><c>True</c> if feature enabled, otherwise returns false.</returns>
		public static bool GetIsFeatureEnabled(UserConnection uc, string code) {
			var featureUtil = ClassFactory.Get<IFeatureUtilities>();
			return featureUtil.GetIsFeatureEnabled(uc, code);
		}

		/// <summary>
		/// Returns is feature disabled for <paramref name="uc"/>.
		/// </summary>
		/// <param name="uc"><see cref="UserConnection"/> instance.</param>
		/// <param name="code">Feature code.</param>
		/// <returns><c>True</c> if feature disabled, otherwise returns false.</returns>
		public static bool GetIsFeatureDisabled(UserConnection uc, string code) {
			return !GetIsFeatureEnabled(uc, code);
		}

		/// <summary>
		/// Calls <paramref name="action"/>. Handles thrown errors.
		/// </summary>
		/// <param name="action"><see cref="Action"/> insntace.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="uc"><see cref="UserConnection"/> instnace.</param>
		public static void TryDoListenerAction(Action action, string senderEmailAddress, UserConnection uc) {
			try {
				action();
			} catch (Exception e) {
				bool isAggregateException = e.GetType() == typeof(AggregateException);
				string exceptionClassName, exceptionMessage;
				if (isAggregateException) {
					exceptionClassName = e.InnerException.GetType().Name;
					exceptionMessage = e.InnerException.Message;
				} else {
					exceptionClassName = e.GetType().Name;
					exceptionMessage = e.Message;
				}
				if (GetIsFeatureDisabled(uc, "OldEmailIntegration")) {
					var helper = ClassFactory.Get<ISynchronizationErrorHelper>
						(new ConstructorArgument("userConnection", uc));
					helper.ProcessSynchronizationError(senderEmailAddress, exceptionClassName, exceptionMessage);
				}
				throw;
			}
		}

		/// <summary>
		/// Returns mailbox last email recived date. When date is older thn synchronization period, then period used.
		/// </summary>
		/// <param name="mailboxId">Mailbox unique identifier.</param>
		/// 
		/// <returns>Mailbox last email recived date.</returns>
		public static DateTime GetFailoverPeriodStartDate(Mailbox mailbox, UserConnection uc) {
			var lastEmailDate = mailbox.GetLastEmailSyncDate(uc);
			var periodDate = mailbox.GetLoadFromDate(uc);
			var date = new List<DateTime> {
				periodDate.Date,
				lastEmailDate
			}.Max();
			return TimeZoneInfo.ConvertTime(date, TimeZoneInfo.Utc, uc.CurrentUser.TimeZone).AddMinutes(FailoverJobFilterMinutesOffset);
		}

		/// <summary>
		/// Returns bpm'online new email events endpoint uri.
		/// </summary>
		/// <returns>Bpm'online new email events endpoint uri.</returns>
		public string GetBpmEndpointUrl() {
			var endpointUrl = SysSettings.GetValue(_uc, "BpmonlineExchangeEventsEndpointUrl", "");
			if (endpointUrl.IsEmpty()) {
				endpointUrl = GetNewBpmonlineExchangeEventsEndpointUrl();
				SysSettings.SetDefValue(_uc, "BpmonlineExchangeEventsEndpointUrl", endpointUrl);
			}
			if (endpointUrl.IsEmpty()) {
				throw new InvalidObjectStateException("Bpmonline exchange events endpoint url cannot be created. " +
					"Fill BpmonlineExchangeEventsEndpointUrl system setting.");
			}
			return endpointUrl;
		}

		#endregion

	}

	#endregion

}
