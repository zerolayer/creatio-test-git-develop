namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Core;

	#region Class: BaseExchangeAttachmentHelper

	/// <summary>
	/// <see cref="IExchangeAttachmentUtilities"/> implementation.
	/// </summary>
	public abstract class BaseExchangeAttachmentHelper : IExchangeAttachmentUtilities
	{
		#region Methods: Public 

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.CreateExchangeService"/>
		/// </summary>
		public abstract ExchangeService CreateExchangeService(UserConnection userConnection, string userEmailAddress);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.SafeBindItem"/>
		/// </summary>
		public abstract EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.SafeBindItem"/>
		/// </summary>
		public abstract EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties, UserConnection userConnection, string senderEmail);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetAttachments"/>
		/// </summary>
		public abstract IEnumerable<Attachment> GetAttachments(EmailMessage message);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetAttachmentsById"/>
		/// </summary>
		public abstract IEnumerable<Attachment> GetAttachmentsById(EmailMessage message, string attachmentId);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.Load"/>
		/// </summary>
		public abstract void Load(Attachment attachment);

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetContent"/>
		/// </summary>
		public abstract byte[] GetContent(Attachment attachment);

		#endregion
	}

	#endregion
}