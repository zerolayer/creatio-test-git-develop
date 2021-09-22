namespace Terrasoft.EmailDomain.Interfaces
{
	using EmailContract.DTO;
	using System;
	using System.Collections.Generic;
	using Terrasoft.EmailDomain.Model;

	#region Interface: IEmailRepository

	/// <summary>
	/// Email model repository interface.
	/// </summary>
	internal interface IEmailRepository
	{

		#region Methods: Internal

		/// <summary>
		/// Saves <paramref name="email"/> to storage.
		/// </summary>
		/// <param name="email"><see cref="EmailModel"/> instnace.</param>
		/// <param name="mailboxId">Mailbox identifier.</param>
		/// <param name="syncSessionId">Synchronization session identifier.</param>
		void Save(EmailModel email, Guid mailboxId = default, string syncSessionId = null);

		/// <summary>
		/// Returns email message headers for <paramref name="messageId"/>.
		/// </summary>
		/// <param name="messageId">Message header identifier.</param>
		/// <returns>Email message headers collection.</returns>
		IEnumerable<EmailModelHeader> GetHeaders(string messageId);

		/// <summary>
		/// Gets <see cref="Email"/> by <paramref name="activityId"/>.
		/// </summary>
		/// <param name="activityId">Activity identifier.</param>
		/// <returns>Email message</returns>
		Email CreateEmail(Guid activityId);

		#endregion

	}

	#endregion

}
