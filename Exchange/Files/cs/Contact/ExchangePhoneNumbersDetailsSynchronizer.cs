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

	#region Class: ExchangePhoneNumbersDetailsSynchronizer

	/// <summary>
	/// Provides methods for Contact phone numbers synchronization with Exchange.
	/// </summary>
	internal class ExchangePhoneNumbersDetailsSynchronizer :
			ExchangeDetailSynchronizer<Exchange.PhoneNumberKey, Exchange.PhoneNumberEntry, Exchange.Contact>
	{

		#region Constructors: Public

		/// <summary>
		/// Creates class instance for contact phone numbers synchronization with Exchange.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localItem">Local storage synchronization element.</param>
		/// <param name="remoteItem">Remote storage synchronization element.</param>
		public ExchangePhoneNumbersDetailsSynchronizer(SyncContext context, LocalItem localItem,
						Exchange.Contact remoteItem)
			: base(context, localItem, "CommunicationTypeId", remoteItem, "ContactCommunication") {
			DetailItems = remoteItem.PhoneNumbers;
			TypesMap = new Dictionary<Exchange.PhoneNumberKey, Guid> {
				{
					Exchange.PhoneNumberKey.BusinessPhone,
					new Guid(CommunicationTypeConsts.WorkPhoneId)
				},
				{
					Exchange.PhoneNumberKey.HomePhone,
					new Guid(CommunicationTypeConsts.HomePhoneId)
				},
				{
					Exchange.PhoneNumberKey.MobilePhone,
					new Guid(CommunicationTypeConsts.MobilePhoneId)
				},
				{
					Exchange.PhoneNumberKey.BusinessPhone2,
					new Guid(CommunicationTypeConsts.WorkPhoneId)
				}
			};
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns exchange phone number value if specific <paramref name="typeKey"/>.
		/// </summary>
		/// <param name="typeKey"><see cref="Exchange.PhoneNumberKey"/> instance.</param>
		/// <returns>
		/// Exchange phone number value.
		/// </returns>
		/// <remarks>
		/// External dependency.
		/// </remarks>
		protected virtual string GetPhoneNumber(Exchange.PhoneNumberKey typeKey) {
			var phone = ExchangeUtility.SafeGetValue<Exchange.PhoneNumberKey, Exchange.PhoneNumberEntry,
							string>(DetailItems, typeKey);
			return phone;
		}

		protected override bool ContainsValue(Exchange.PhoneNumberKey typeKey) {
			var phone = GetPhoneNumber(typeKey);
			return !string.IsNullOrEmpty(phone);
		}

		protected override void SetLocalItemValue(Entity detailItem, Exchange.PhoneNumberKey typeKey) {
			var phone = GetPhoneNumber(typeKey);
			if (string.IsNullOrEmpty(phone)) {
				return;
			}
			detailItem.SetColumnValue("Number", phone);
			detailItem.SetColumnValue(DetailItemTypeColumnName, TypesMap[typeKey]);
		}

		protected override void SetRemoteItemValue(Entity detailItem, Exchange.PhoneNumberKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			remoteContact.PhoneNumbers[typeKey] = detailItem.GetTypedColumnValue<string>("Number");
		}

		protected override void DeleteRemoteDetail(Exchange.PhoneNumberKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			remoteContact.PhoneNumbers[typeKey] = null;
		}

		#endregion

	}

	#endregion
}
