namespace Terrasoft.IntegrationV2.Logging
{
	using System;
	using global::Common.Logging;
	using Newtonsoft.Json;
	using Terrasoft.Core.Factories;
	using Terrasoft.IntegrationV2.Logging.Interfaces;
	using Terrasoft.Messaging.Common;

	#region Class: SynchronizationLogger

	[DefaultBinding(typeof(ISynchronizationLogger))]
	internal class SynchronizationLogger : ISynchronizationLogger
	{

		#region Fields: Private

		private readonly ILog _log = LogManager.GetLogger("ExchangeListener");

		private readonly Guid _userId;

		#endregion

		#region Constructors: Public

		public SynchronizationLogger(Guid userId) {
			_userId = userId;
		}

		#endregion

		#region Methods: Private

		private void SendMsgToClient(string msg, string bodyType = "Info", Exception e = null) {
			try {
				var channel = MsgChannelManager.Instance.FindItemByUId(_userId);
				if (channel == null) {
					return;
				}
				var messageBody = e == null
					? msg
					: string.Concat(msg, "\r\n", e.ToString());
				var simpleMessage = new SimpleMessage {
					Id = _userId,
					Body = JsonConvert.SerializeObject(messageBody),
				};
				simpleMessage.Header.Sender = "SyncMsgLogger";
				simpleMessage.Header.BodyTypeName = bodyType;
				channel.PostMessage(simpleMessage);
			}
			catch (Exception ex) {
				_log?.Error($"Socket channel not oppened for {_userId}", ex);
			}
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="ISynchronizationLogger.ErrorFormat(string, Exception, object[])"/>
		public void ErrorFormat(string format, Exception exception = null, params object[] args) {
			SendMsgToClient(string.Format(format, args), "Error", exception);
			_log.ErrorFormat(format, exception, args);
		}

		/// <inheritdoc cref="ISynchronizationLogger.InfoFormat(string, object[])"/>
		public void InfoFormat(string format, params object[] args) {
			SendMsgToClient(string.Format(format, args));
			_log.InfoFormat(format, args);
		}

		/// <inheritdoc cref="ISynchronizationLogger.DebugFormat(string, object[])"/>
		public void DebugFormat(string format, params object[] args) {
			SendMsgToClient(string.Format(format, args));
			_log.DebugFormat(format, args);
		}

		/// <inheritdoc cref="ISynchronizationLogger.Error(string, Exception)"/>
		public void Error(string message, Exception exception = null) {
			SendMsgToClient(message, "Error", exception);
			_log.Error(message, exception);
		}

		/// <inheritdoc cref="ISynchronizationLogger.Warn(string)"/>
		public void Warn(string message) {
			SendMsgToClient(message);
			_log.Warn(message);
		}

		/// <inheritdoc cref="ISynchronizationLogger.Info(string)"/>
		public void Info(string message) {
			SendMsgToClient(message);
			_log.Info(message);
		}

		#endregion

	}

	#endregion

}
