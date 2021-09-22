namespace IntegrationV2.MailboxDomain
{
	using System;
	using System.Linq;
	using IntegrationApi.MailboxDomain.Interfaces;
	using IntegrationApi.MailboxDomain.Model;
	using IntegrationV2.MailboxDomain.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;

	#region Class: MailboxService

	/// <summary>
	/// Mailbox service implementation.
	/// </summary>
	[DefaultBinding(typeof(IMailServerService))]
	public class MailServerService : IMailServerService
	{

		#region Fields: Private

		private readonly IMailServerRepository _mailServerRepository;

		#endregion

		#region Constructors: Public

		public MailServerService(UserConnection uc) {
			_mailServerRepository = ClassFactory.Get<IMailServerRepository>(new ConstructorArgument("uc", uc));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IMailServerService.GetServer(Guid, bool)"/>
		public MailServer GetServer(Guid id, bool useForSynchronization = true) {
			return _mailServerRepository.GetAll(useForSynchronization)
				.FirstOrDefault(m => m.Id.Equals(id));
		}

		#endregion

	}

	#endregion

}
