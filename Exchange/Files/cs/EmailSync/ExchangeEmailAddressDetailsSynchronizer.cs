namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Sync;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeEmailAddressDetailsSynchronizer

	/// <summary>
	/// Provides methods for Contact email address synchronization with Exchange.
	/// </summary>
	internal class ExchangeEmailAddressDetailsSynchronizer :
			ExchangeDetailSynchronizer<Exchange.EmailAddressKey, Exchange.EmailAddressEntry, Exchange.Contact>
	{

		#region Constructors: Public

		/// <summary>
		/// Creates class instance for contact email address synchronization with Exchange.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localItem">Local storage synchronization element.</param>
		/// <param name="remoteItem">Remote storage synchronization element.</param>
		public ExchangeEmailAddressDetailsSynchronizer(SyncContext context, LocalItem localItem,
						Exchange.Contact remoteItem)
			: base(context, localItem, "CommunicationTypeId", remoteItem, "ContactCommunication") {
			DetailItems = remoteItem.EmailAddresses;
			TypesMap = new Dictionary<Exchange.EmailAddressKey, Guid> {
				{
					Exchange.EmailAddressKey.EmailAddress1,
					new Guid(CommunicationTypeConsts.EmailId)
				},
				{
					Exchange.EmailAddressKey.EmailAddress2,
					new Guid(CommunicationTypeConsts.EmailId)
				},
				{
					Exchange.EmailAddressKey.EmailAddress3,
					new Guid(CommunicationTypeConsts.EmailId)
				}
			};
		}

		#endregion

		#region Methods: Protected

		protected override void SetLocalItemValue(Entity detailItem, Exchange.EmailAddressKey typeKey) {
			Exchange.EmailAddress emailAddress = ExchangeUtility.SafeGetValue<Exchange.EmailAddressKey,
					Exchange.EmailAddressEntry, Exchange.EmailAddress>(DetailItems, typeKey);
			if (emailAddress == null) {
				return;
			}
			detailItem.SetColumnValue("Number", emailAddress.Address);
			detailItem.SetColumnValue(DetailItemTypeColumnName, TypesMap[typeKey]);
		}

		protected override void SetRemoteItemValue(Entity detailItem, Exchange.EmailAddressKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			remoteContact.EmailAddresses[typeKey] = detailItem.GetTypedColumnValue<string>("Number");
		}

		protected override void DeleteRemoteDetail(Exchange.EmailAddressKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			var remoteProvider = Context.RemoteProvider as ExchangeContactSyncProviderImpl;
			if (remoteProvider == null) {
				Context.LogError(SyncAction.Delete, SyncDirection.Upload,
					"[ExchangeEmailAddressDetailsSynchronizer.DeleteRemoteDetail]: RemoteProvider type mismatch! " +
					"Should be of type {0}. Remote item Id: {1}", typeof(ExchangeContactSyncProviderImpl),
					remoteContact.Id.UniqueId);
				return;
			}
			remoteProvider.DeleteContactEmailAddress(Context, remoteContact.Id.UniqueId, typeKey);
		}

		protected override bool ContainsValue(Exchange.EmailAddressKey typeKey) {
			Exchange.EmailAddress emailAddress = ExchangeUtility.SafeGetValue<Exchange.EmailAddressKey,
					Exchange.EmailAddressEntry, Exchange.EmailAddress>(DetailItems, typeKey);
			if (emailAddress == null) {
				return false;
			}
			return !string.IsNullOrEmpty(emailAddress.Address);
		}

		#endregion

	}

	#endregion

}
