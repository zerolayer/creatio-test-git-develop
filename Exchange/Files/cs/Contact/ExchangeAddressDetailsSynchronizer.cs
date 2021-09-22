namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Sync;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeAddressDetailsSynchronizer

	/// <summary>
	/// Provides methods for Contact address synchronization with Exchange.
	/// </summary>
	public class ExchangeAddressDetailsSynchronizer :
			ExchangeDetailSynchronizer<Exchange.PhysicalAddressKey, Exchange.PhysicalAddressEntry, Exchange.Contact>
	{

		#region Constructors: Public

		/// <summary>
		/// Creates class instance for contact address synchronization with Exchange.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localItem">Local storage synchronization element.</param>
		/// <param name="remoteItem">Remote storage synchronization element.</param>
		public ExchangeAddressDetailsSynchronizer(SyncContext context, LocalItem localItem,
						Exchange.Contact remoteItem)
			: base(context, localItem, "AddressTypeId", remoteItem, "ContactAddress") {
			if (remoteItem != null) {
				DetailItems = remoteItem.PhysicalAddresses;
			}
			TypesMap = AddressTypesMap;
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// <see cref="Guid"/> values of type address from <see cref="Exchange.PhysicalAddressKey"/> enumeration.
		/// </summary>
		public static Dictionary<Exchange.PhysicalAddressKey, Guid> AddressTypesMap =
					new Dictionary<Exchange.PhysicalAddressKey, Guid> {
						{
							Exchange.PhysicalAddressKey.Home, ExchangeConsts.HomeAddressTypeId
						},
						{
							Exchange.PhysicalAddressKey.Business, ExchangeConsts.BusinessAddressTypeId
						}
					};

			#endregion

		#region Methods: Private

		private void AppendToAddressString(StringBuilder address, string value) {
			if (string.IsNullOrEmpty(value)) {
				return;
			}
			if (address.Length > 0) {
				address.Append(", " + value);
			} else {
				address.Append(value);
			}

		}

		#endregion

		#region Methods: Protected

		protected override bool ContainsValue(Exchange.PhysicalAddressKey typeKey) {
			Exchange.PhysicalAddressEntry address = ExchangeUtility.SafeGetValue<Exchange.PhysicalAddressKey,
					Exchange.PhysicalAddressEntry, Exchange.PhysicalAddressEntry>(DetailItems, typeKey);
			if (address == null) {
				return false;
			}
			return !string.IsNullOrEmpty(address.City + address.CountryOrRegion + address.State + address.Street);
		}

		protected override void SetLocalItemValue(Entity detailItem, Exchange.PhysicalAddressKey typeKey) {
			var localAddr = detailItem as ContactAddress;
			if (localAddr == null) {
				return;
			}
			var contactProvider = Context.RemoteProvider as ExchangeContactSyncProviderImpl;
			if (contactProvider == null) {
				return;
			}
			Exchange.PhysicalAddressEntry exchangeAddr = ExchangeUtility.SafeGetValue<Exchange.PhysicalAddressKey,
					Exchange.PhysicalAddressEntry, Exchange.PhysicalAddressEntry>(DetailItems, typeKey);
			if (exchangeAddr == null) {
				return;
			}
			Dictionary<string, AddressDetail> addressesLookupMap = (contactProvider).AddressesLookupMap;
			if (addressesLookupMap == null) {
				return;
			}
			var address = new StringBuilder();
			AppendToAddressString(address, exchangeAddr.Street);
			localAddr.Zip = exchangeAddr.PostalCode;
			localAddr.AddressTypeId = TypesMap[typeKey];
			if (!addressesLookupMap.Any()) {
				localAddr.Address = address.ToString();
				return;
			}
			string cityName = exchangeAddr.City;
			string regionName = exchangeAddr.State;
			string countryName = exchangeAddr.CountryOrRegion;
			string addressKey = ExchangeContactAddressDetailHelper.GetUniqueKey(exchangeAddr);
			AddressDetail mapItem = default(AddressDetail);
			if (addressesLookupMap.Keys.Contains(addressKey)) {
				mapItem = addressesLookupMap[addressKey];
			}
			if (mapItem.CityId != Guid.Empty) {
				localAddr.CityId = mapItem.CityId;
			} else {
				localAddr.SetColumnValue("CityId", null);
				AppendToAddressString(address, cityName);
			}
			if (mapItem.RegionId != Guid.Empty) {
				localAddr.RegionId = mapItem.RegionId;
			} else {
				localAddr.SetColumnValue("RegionId", null);
				AppendToAddressString(address, regionName);
			}
			if (mapItem.CountryId != Guid.Empty) {
				localAddr.CountryId = mapItem.CountryId;
			} else {
				localAddr.SetColumnValue("CountryId", null);
				AppendToAddressString(address, countryName);
			}
			localAddr.Address = address.ToString();
		}

		protected override void SetRemoteItemValue(Entity detailItem, Exchange.PhysicalAddressKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			var localAddress = detailItem as ContactAddress;
			if (localAddress == null) {
				return;
			}
			var remoteAddress = new Exchange.PhysicalAddressEntry {
				City = localAddress.CityName,
				CountryOrRegion = localAddress.CountryName,
				State = localAddress.RegionName,
				PostalCode = localAddress.Zip,
				Street = localAddress.Address
			};
			remoteContact.PhysicalAddresses[typeKey] = null;
			remoteContact.PhysicalAddresses[typeKey] = remoteAddress;
		}

		protected override void DeleteRemoteDetail(Exchange.PhysicalAddressKey typeKey) {
			var remoteContact = RemoteItem;
			if (remoteContact == null) {
				return;
			}
			remoteContact.PhysicalAddresses[typeKey] = null;
		}

		#endregion

	}

	#endregion
}