namespace IntegrationV2.MailboxDomain.Interfaces
{
	using System;
	using System.Collections.Generic;
	using IntegrationApi.MailboxDomain.Model;

	#region Interface: IMailServerRepository

	/// <summary>
	/// Mailbox provider storage repository interface.
	/// </summary>
	internal interface IMailServerRepository
	{

		#region Methods: Internal

		/// <summary>
		/// Returns all mailbox providers list.
		/// </summary>
		/// <param name="useForSynchronization">Sign is synchronization mode or not.</param>
		/// <returns><see cref="MailServer"/> collection.</returns>
		IEnumerable<MailServer> GetAll(bool useForSynchronization = true);

		/// <summary>
		/// Returns concrete mailbox provider.
		/// </summary>
		/// <param name="mailServerId">Mailserver Id.</param>
		/// <returns><see cref="MailServer"/> instance.</returns>
		MailServer GetById(Guid mailServerId);

		#endregion

	}

	#endregion

}
