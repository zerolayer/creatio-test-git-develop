namespace Terrasoft.IntegrationV2.Logging.Interfaces
{
	using System;

	#region Interface: ISynchronizationLogger

	/// <summary>
	/// Writes logs to standart logger and sends message to client.
	/// </summary>
	internal interface ISynchronizationLogger
	{

		#region Methods: Internal

		/// <inheritdoc cref="global::Common.Logging.ILog.ErrorFormat(string, Exception, object[])"/>
		void ErrorFormat(string format, Exception exception = null, params object[] args);

		/// <inheritdoc cref="global::Common.Logging.ILog.InfoFormat(string, object[])"/>
		void InfoFormat(string format, params object[] args);

		/// <inheritdoc cref="global::Common.Logging.ILog.DebugFormat(string, object[])"/>
		void DebugFormat(string format, params object[] args);

		/// <inheritdoc cref="global::Common.Logging.ILog.Error(string, Exception)"/>
		void Error(string message, Exception exception = null);

		/// <inheritdoc cref="global::Common.Logging.ILog.Warn(string)"/>
		void Warn(string message);

		/// <inheritdoc cref="global::Common.Logging.ILog.Info(string)"/>
		void Info(string message);

		#endregion

	}

	#endregion

}
