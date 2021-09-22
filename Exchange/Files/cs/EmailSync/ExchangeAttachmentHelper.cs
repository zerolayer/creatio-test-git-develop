namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using System.Linq;
	using Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Core;
    using Terrasoft.Core.Factories;

	#region Class: ExchangeAttachmentHelper

	/// <summary>
	/// <see cref="IExchangeAttachmentUtilities"/> implementation.
	/// </summary>
	[DefaultBinding(typeof(IExchangeAttachmentUtilities), Name = "JobUserConnection")]
	[DefaultBinding(typeof(IExchangeAttachmentUtilities), Name = "UserConnection")]
	[DefaultBinding(typeof(IExchangeAttachmentUtilities), Name = "ActorUserConnection")]
	public class ExchangeAttachmentHelper : BaseExchangeAttachmentHelper
	{
		#region Methods: Public 

		/// <summary>
		/// <see cref="IExchangeAttachmentHelper.CreateExchangeService"/>
		/// </summary>
		public override ExchangeService CreateExchangeService(UserConnection userConnection, string userEmailAddress) {
			return new ExchangeUtilityImpl().CreateExchangeService(userConnection, userEmailAddress);
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.SafeBindItem"/>
		/// </summary>
		public override EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties) {
			return new ExchangeUtilityImpl().SafeBindItem<EmailMessage>(service, itemId, properties);
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.SafeBindItem"/>
		/// </summary>
		public override EmailMessage SafeBindItem(ExchangeService service, ItemId itemId, PropertySet properties, UserConnection userConnection, string senderEmail) {
			return new ExchangeUtilityImpl().SafeBindItem<EmailMessage>(service, itemId, properties);
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetAttachments"/>
		/// </summary>
		public override IEnumerable<Attachment> GetAttachments(EmailMessage message) {
			return message.Attachments;
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetAttachmentsById"/>
		/// </summary>
		public override IEnumerable<Attachment> GetAttachmentsById(EmailMessage message, string attachmentId) {
			return GetAttachments(message).Where(e => e.Id == attachmentId);
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.Load"/>
		/// </summary>
		public override void Load(Attachment attachment) {
			attachment.Load();
		}

		/// <summary>
		/// <see cref="IExchangeAttachmentUtilities.GetContent"/>
		/// </summary>
		public override byte[] GetContent(Attachment attachment) {
			var fileAttachment = attachment as FileAttachment;
			if (fileAttachment != null) {
				return fileAttachment.Content;
				
			}
			var itemAttachment = attachment as ItemAttachment;
			if (itemAttachment != null) {
				var propertySet = new PropertySet(BasePropertySet.IdOnly);
				propertySet.Add(ItemSchema.Attachments);
				propertySet.Add(ItemSchema.MimeContent);
				itemAttachment.Load(propertySet);
				var mimeContent = itemAttachment.Item.MimeContent;
				return mimeContent.Content;

			}
			return new byte[0];
		}

		#endregion
	}

	#endregion
}
