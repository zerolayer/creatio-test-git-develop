namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Core;

	#region Interface: IExchangeAttachmentUtilities
	
	/// <summary>
	/// Exchange attachments utilities interface.
	/// External dependency wraper.
	/// </summary>
	public interface IExchangeAttachmentUtilities
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
		/// Calls bind methods for <see cref="EmailMessage"/> with <paramref name="itemId"/>.
		/// </summary>
		/// <param name="service"><see cref="ExchangeService"/> instance.</param>
		/// <param name="itemId"><see cref="ItemId"/> instance.</param>
		/// <param name="properties"><see cref="PropertySet"/> instance.</param>
		/// <returns>Bound <see cref="EmailMessage"/> instance.</returns>
		EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties);

		/// <summary>
		/// Calls bind methods for <see cref="EmailMessage"/> with <paramref name="itemId"/>.
		/// </summary>
		/// <param name="service"><see cref="ExchangeService"/> instance.</param>
		/// <param name="itemId"><see cref="ItemId"/> instance.</param>
		/// <param name="properties"><see cref="PropertySet"/> instance.</param>
		/// <param name="userConnection"><see cref="Terrasoft.Core.UserConnection"/> instance.</param>
		/// <param name="senderEmail">Synchronization email.</param>
		/// <returns>Bound <see cref="EmailMessage"/> instance.</returns>
		EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties, UserConnection userConnection, string senderEmail);

		/// <summary>
		/// Returns <paramref name="message"/> attachments collection.
		/// </summary>
		/// <param name="message"><see cref="EmailMessage"/> instance.</param>
		/// <returns><paramref name="message"/> attachments collection.</returns>
		IEnumerable<Attachment> GetAttachments(EmailMessage message);

		/// <summary>
		/// Returns <paramref name="message"/> attachments collection filtered by <paramref name="attachmentId"/>.
		/// </summary>
		/// <param name="message"><see cref="EmailMessage"/> instance.</param>
		/// <param name="attachmentId">Attachment id.</param>
		/// <returns>Filtered <paramref name="message"/> attachments collection.</returns>
		IEnumerable<Attachment> GetAttachmentsById(EmailMessage message, string attachmentId);

		/// <summary>
		/// Calls <see cref="Attachment.Load()"/> method.
		/// </summary>
		/// <param name="attachment"><see cref="Attachment"/> instance.</param>
		void Load(Attachment attachment);

		/// <summary>
		/// Returns <paramref name="attachment"/> content.
		/// </summary>
		/// <param name="attachment"><see cref="Attachment"/> instance.</param>
		/// <returns><paramref name="attachment"/> content.</returns>
		byte[] GetContent(Attachment attachment);

		#endregion

	}

	#endregion

}