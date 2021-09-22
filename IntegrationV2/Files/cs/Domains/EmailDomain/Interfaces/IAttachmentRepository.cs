namespace Terrasoft.EmailDomain.Interfaces
{
	using EmailContract.DTO;
	using System;
	using System.Collections.Generic;
	using Terrasoft.EmailDomain.Model;

	#region Interface: IAttachmentRepository

	/// <summary>
	/// Attachment model repository interface.
	/// </summary>
	internal interface IAttachmentRepository
	{

		#region Methods: Internal

		/// <summary>
		/// Saves attachments from <paramref name="email"/> to storage.
		/// </summary>
		/// <param name="email"><see cref="EmailModel"/> instance.</param>
		void SaveAttachments(EmailModel email);

		/// <summary>
		/// Returns file service link for <paramref name="attachmentId"/>.
		/// </summary>
		/// <param name="attachmentId">Attachment identifier.</param>
		/// <returns>Attachment file service link.</returns>
		string GetAttachmentLink(Guid attachmentId);

		/// <summary>
		/// Set inline flag.
		/// </summary>
		/// <param name="attachmentId">Attachment identifier.</param>
		void SetInline(Guid attachmentId);

		/// <summary>
		/// Get email attachments for <paramref name="activityId"/>.
		/// </summary>
		/// <param name="activityId"></param>
		/// <returns>Email attachment collection.</returns>
		List<Attachment> GetAttachments(Guid activityId);

		#endregion

	}

	#endregion

}
