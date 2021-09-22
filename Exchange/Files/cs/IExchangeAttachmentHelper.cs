namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Core;

	#region Interface: IExchangeAttachmentHelper
	
	/// <summary>
	/// Exchange attachments helper interface.
	/// External dependency wraper.
	/// </summary>
	public interface IExchangeAttachmentHelper
	{

		#region Methods: Public

		/// <summary>
		/// Creates <see cref="Exchange.ExchangeService"/> instance.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="userEmailAddress">Sender email address.</param>
		/// <returns><see cref="Exchange.ExchangeService"/> instance.</returns>
		ExchangeService CreateExchangeService(UserConnection userConnection, string userEmailAddress);

		/// <summary>
		/// Calls bind methods for <see cref="Exchange.EmailMessage"/> with <paramref name="itemId"/>.
		/// </summary>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="itemId"><see cref="Exchange.ItemId"/> instance.</param>
		/// <param name="properties"><see cref="Exchange.PropertySet"/> instance.</param>
		/// <returns>Bound <see cref="Exchange.EmailMessage"/> instance.</returns>
		EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties);

		/// <summary>
		/// Returns <paramref name="message"/> attachments collection.
		/// </summary>
		/// <param name="message"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <returns><paramref name="message"/> attachments collection.</returns>
		IEnumerable<Attachment> GetEmailAttachments(EmailMessage message);

		/// <summary>
		/// Returns <paramref name="message"/> attachments collection filtered by <paramref name="attachmentId"/>.
		/// </summary>
		/// <param name="message"><see cref="Exchange.EmailMessage"/> instance.</param>
		/// <param name="attachmentId">Attachment id.</param>
		/// <returns>Filtered <paramref name="message"/> attachments collection.</returns>
		IEnumerable<Attachment> GetEmailAttachmentsById(EmailMessage message, string attachmentId);

		/// <summary>
		/// Calls <see cref="Exchange.FileAttachment.Load()"/> method.
		/// </summary>
		/// <param name="attachment"><see cref="Exchange.FileAttachment"/> instance.</param>
		void LoadAttachment(FileAttachment attachment);

		/// <summary>
		/// Returns <paramref name="attachment"/> content.
		/// </summary>
		/// <param name="attachment"><see cref="Exchange.FileAttachment"/> instance.</param>
		/// <returns><paramref name="attachment"/> content.</returns>
		byte[] GetAttachmentContent(FileAttachment attachment);

		#endregion

	}

	#endregion

}