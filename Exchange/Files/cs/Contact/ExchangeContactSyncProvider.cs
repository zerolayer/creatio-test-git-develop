namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using System.Data;
	using Terrasoft.Configuration;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Common.Json;
	using Terrasoft.Common;
	using Newtonsoft.Json.Linq;
	using Exchange = Microsoft.Exchange.WebServices.Data;
	using Terrasoft.Nui.ServiceModel.Extensions;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeContactSyncProviderImpl

	/// <summary>
	/// Class for synchronization with Exchange contacts.
	/// </summary>
	[DefaultBinding(typeof(BaseExchangeSyncProvider), Name = "ExchangeContactSyncProvider")]
	public class ExchangeContactSyncProviderImpl : ExchangeSyncProvider
	{

		#region Constants: Private

		private const int EsqParamsCount = 1000;

		#endregion

		#region Fields: Private

		private static readonly Dictionary<Type, Type> SyncTypeToTypeMap = new Dictionary<Type, Type> {
			{ typeof(ExchangeContact), typeof(Microsoft.Exchange.WebServices.Data.Contact) }
		};

		private readonly ContactEmailAddressPropertiesMap _emailAddressPropertiesMap =
			new ContactEmailAddressPropertiesMap();

		/// <summary>
		/// <see cref="Terrasoft.Core.UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		/// <summary>
		/// <see cref="Terrasoft.Configuration.SynchronizationErrorHelper"/> instance.
		/// </summary>
		private readonly SynchronizationErrorHelper _syncErrorHelper;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// Initialize new instance of <see cref="ExchangeContactSyncProviderImpl" /> with passed <paramref name="settings"/>.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="settings"><see cref="ContactExchangeSettings"/> instance.</param>
		public ExchangeContactSyncProviderImpl(UserConnection userConnection, string senderEmailAddress,
				ContactExchangeSettings settings = null) : 
				base(ExchangeConsts.ExchangeContactStoreId, userConnection.CurrentUser.TimeZone, senderEmailAddress) {
			_userConnection = userConnection;
			UserSettings = settings ?? new ContactExchangeSettings(userConnection, senderEmailAddress);
			Version = UserSettings.LastSyncDate;
			AccountsMap = new Dictionary<string, Guid>();
			_syncErrorHelper = SynchronizationErrorHelper.GetInstance(userConnection);
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// User synchronization settings.
		/// </summary>
		public ContactExchangeSettings UserSettings {
			get;
			private set;
		}

		/// <summary>
		/// <see cref="Account"/> instances ids, selected by name from <see cref="Exchange.Contact"/>.
		/// </summary>
		public Dictionary<string, Guid> AccountsMap {
			get;
			private set;
		}

		/// <summary>
		/// <see cref="AddressDetail"/> instances, selected by country, region and city name.
		/// </summary>
		public Dictionary<string, AddressDetail> AddressesLookupMap {
			get;
			private set;
		}

		#endregion

		#region Methods: Private

		private static IEnumerable<Entity> GetDetailsDataByFilter(
						UserConnection userConnection, DetailEntityConfig detailConfig, Guid filterValue) {
			var esq = new EntitySchemaQuery(userConnection.EntitySchemaManager.FindInstanceByName(
					detailConfig.SchemaName
				));
			esq.AddAllSchemaColumns();
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal,
					detailConfig.ForeingKeyColumnName, filterValue));
			return esq.GetEntityCollection(userConnection);
		}

		private void AddAddressDetails(ICollection<ExchangeAddressDetail> exchangeAddrDetails,
						Exchange.Contact contact) {
			foreach (KeyValuePair<Exchange.PhysicalAddressKey, Guid> addresstypeMap in
							ExchangeAddressDetailsSynchronizer.AddressTypesMap) {
				Exchange.PhysicalAddressEntry address = ExchangeUtility.SafeGetValue<Exchange.PhysicalAddressKey,
						Exchange.PhysicalAddressEntry, Exchange.PhysicalAddressEntry>(contact.PhysicalAddresses,
								addresstypeMap.Key);
				if (address.IsEmpty()) {
					exchangeAddrDetails.Add(new ExchangeAddressDetail(address.CountryOrRegion ?? "",
							address.State ?? "", address.City ?? ""));
				}
			}
		}

		private IEnumerable<LocalItem> GetNotSyncedContacts(SyncContext context, SyncItemSchema primarySchema) {
			string primarySchemaName = primarySchema.PrimaryEntityConfig.SchemaName;
			UserConnection userConnection = context.UserConnection;
			DateTime lastSyncVersion = context.LastSyncVersion;
			var esq = new EntitySchemaQuery(userConnection.EntitySchemaManager, primarySchemaName);
			esq.PrimaryQueryColumn.IsAlwaysSelect = true;
			LocalProvider.AddItemSchemaColumns(esq, primarySchema.PrimaryEntityConfig);
			AddExportFiltersBySettings(context, esq);
			Select select = esq.GetSelectQuery(context.UserConnection);
			select.And("Contact", "Id").Not()
					.Exists(new Select(userConnection)
									.Column("LocalId")
									.From("SysSyncMetaData")
									.Where("SysSyncMetaData", "CreatedById").IsEqual(new QueryParameter(
											userConnection.CurrentUser.ContactId))
									.And("SysSyncMetaData", "RemoteItemName").IsEqual(new QueryParameter(
											"ExchangeContact"))
									.And("SysSyncMetaData", "SyncSchemaName").IsEqual(new QueryParameter(
											"Contact"))
									.And("Contact", "Id").IsEqual("SysSyncMetaData", "LocalId"));
			EntityCollection entities = esq.GetEntityCollection(context.UserConnection);
			foreach (Entity entity in entities) {
				var localItem = new LocalItem(primarySchema);
				localItem.Entities[primarySchemaName].Add(new SyncEntity(entity, SyncState.New));
				foreach (DetailEntityConfig detailConfig in primarySchema.DetailConfigs.Where(d =>
						d.PrimarySchemaName == primarySchemaName)) {
					IEnumerable<SyncEntity> details = GetDetailsDataByFilter(context.UserConnection, detailConfig,
							entity.PrimaryColumnValue).Select(SyncEntity.CreateNew);
					localItem.Entities[detailConfig.SchemaName].AddRange(details);
				}
				yield return localItem;
			}
		}

		private void AddExportFiltersBySettings(SyncContext context, EntitySchemaQuery esq) {
			UserConnection userConnection = context.UserConnection;
			esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "CreatedBy",
					userConnection.CurrentUser.ContactId));
			if (UserSettings.ExportContactsAll) {
				return;
			}
			var contactTypeFilters = new EntitySchemaQueryFilterCollection(esq, LogicalOperationStrict.Or);
			if (UserSettings.ExportContactsEmployers) {
				contactTypeFilters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "Type",
				ExchangeConsts.ContactEmployeeTypeId));
			}
			if (UserSettings.ExportContactsCustomers) {
				contactTypeFilters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "Type",
				ExchangeConsts.ContactCustomerTypeId));
			}
			if (UserSettings.ExportContactsOwner) {
				contactTypeFilters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false, "Owner",
				userConnection.CurrentUser.ContactId));
			}
			if (UserSettings.ExportContactsFromGroups) {
				AddContactsDynamicGroupFilters(UserSettings.LocalFolderUIds, userConnection, contactTypeFilters);
			}
			if (!contactTypeFilters.Any()) {
				esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, false,
					esq.RootSchema.PrimaryColumn.Name, Guid.Empty));
				return;
			}
			esq.Filters.Add(contactTypeFilters);
		}

		private void AddContactsDynamicGroupFilters(
						IDictionary<string, Guid> localFolderUIds, UserConnection userConnection,
						EntitySchemaQueryFilterCollection filters) {
			if (!localFolderUIds.Any()) {
				return;
			}
			EntitySchemaManager entitySchemaManager = userConnection.EntitySchemaManager;
			var foldersEsq = new EntitySchemaQuery(entitySchemaManager, "ContactFolder");
			string searchDataColumnName = foldersEsq.AddColumn("SearchData").Name;
			string[] folderIdsStrArray =
								(from folderId in localFolderUIds.Values
								 select folderId.ToString()).ToArray();
			foldersEsq.Filters.Add(foldersEsq.CreateFilterWithParameters(FilterComparisonType.Equal, false,
					"Id", folderIdsStrArray));
			EntityCollection folderEntities = foldersEsq.GetEntityCollection(userConnection);
			foreach (Entity folderEntity in folderEntities) {
				byte[] data = folderEntity.GetBytesValue(searchDataColumnName);
				string serializedFilters = Encoding.UTF8.GetString(data, 0, data.Length);
				EntitySchema entitySchema = entitySchemaManager.GetInstanceByName("Contact");
				var dataSourceFilters =
										Json.Deserialize<Terrasoft.Nui.ServiceModel.DataContract.Filters>(serializedFilters);
				IEntitySchemaQueryFilterItem esqFilters =
										dataSourceFilters.BuildEsqFilter(entitySchema.UId, userConnection);
				if (esqFilters != null) {
					filters.Add(esqFilters);
				}
			}
		}

		/// <summary>
		/// Updates last sync date value.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		private void UpdateLastSyncDate(UserConnection userConnection, SyncContext context) {
			var update = new Update(userConnection, "ContactSyncSettings")
								.Set("LastSyncDate", Column.Parameter(context.CurrentSyncStartVersion))
								.Where("MailboxSyncSettingsId")
								.IsEqual(new Select(userConnection).Top(1)
									.Column("Id")
									.From("MailboxSyncSettings")
									.Where("SenderEmailAddress")
									.IsEqual(Column.Parameter(UserSettings.SenderEmailAddress)));
			update.Execute();
		}

		/// <summary>
		/// Fills exchange companies to bpmonline account dictionary.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="exchangeCompanies">Exchange company names.</param>
		private void FillAccountsMap(SyncContext context, List<string> exchangeCompanies) {
			if (exchangeCompanies == null || !exchangeCompanies.Any()) {
				return;
			}
			if (UserSettings.LinkContactToAccountType == ExchangeConsts.NeverLinkContactToAccountId) {
				return;
			}
			IEnumerable<string> distinctExchangeCompanies = exchangeCompanies.Distinct();
			int index = 0;
			while (index < distinctExchangeCompanies.Count()) {
				IEnumerable<string> paramsValue = distinctExchangeCompanies.Skip(index).Take(EsqParamsCount);
				index += EsqParamsCount;
				var accountsEsq = new EntitySchemaQuery(context.UserConnection.EntitySchemaManager, "Account");
				accountsEsq.PrimaryQueryColumn.IsAlwaysSelect = true;
				accountsEsq.AddColumn("Name");
				accountsEsq.Filters.Add(accountsEsq.CreateFilterWithParameters(FilterComparisonType.Equal, false,
						"Name", paramsValue.ToArray()));
				EntityCollection accounts = accountsEsq.GetEntityCollection(context.UserConnection);
				if (accounts.Any()) {
					foreach (Entity account in accounts) {
						var accountName = account.GetTypedColumnValue<string>("Name");
						if (!AccountsMap.ContainsKey(accountName)) {
							AccountsMap[accountName] = account.PrimaryColumnValue;
						}
					}
				}
			}
			IEnumerable<string> notFoundCompanies = from company in exchangeCompanies
													where AccountsMap.All(am => am.Key != company)
													select company;
			var companies = notFoundCompanies as IList<string> ?? notFoundCompanies.ToList();
			if (!companies.Any() ||
					UserSettings.LinkContactToAccountType != ExchangeConsts.AllwaysLinkContactToAccountId)
				return;
			foreach (string company in companies) {
				EntitySchema accountSchema =
									context.UserConnection.EntitySchemaManager.FindInstanceByName("Account");
				Entity account = accountSchema.CreateEntity(context.UserConnection);
				account.SetDefColumnValues();
				account.SetColumnValue("Name", company);
				AccountsMap[company] = account.GetTypedColumnValue<Guid>("Id");
				account.Save(false);
			}
		}

		/// <summary>
		///  Fills exchange address to bpmonline address dictionary.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="exchangeAddrValues">Exchange address details.</param>
		private void FillAddressDetailsData(SyncContext context, List<ExchangeAddressDetail> exchangeAddrValues) {
			if (exchangeAddrValues.Count == 0) {
				return;
			}
			var addressDetailHelper = new ExchangeContactAddressDetailHelper(context.UserConnection,
				exchangeAddrValues);
			AddressesLookupMap = addressDetailHelper.GetAddressesLookupMap();
		}

		#endregion

		#region Methods:Public

		public Exchange.SearchFilter GetContactsFilters() {
			if (UserSettings.LastSyncDate != DateTime.MinValue) {
				DateTime lastSyncDateUtc =
					TimeZoneInfo.ConvertTimeToUtc(UserSettings.LastSyncDate, TimeZone);
				var lastSyncDateUtcFilter = new Exchange.SearchFilter.IsGreaterThan(
						Exchange.ItemSchema.LastModifiedTime, lastSyncDateUtc.ToLocalTime());
				var filterCollection = new Exchange.SearchFilter.SearchFilterCollection(Exchange.LogicalOperator.Or);
				var customPropSetFilter = new Exchange.SearchFilter.Exists(ExchangeUtilityImpl.LocalIdProperty);
				var notCustomPropSetFilter = new Exchange.SearchFilter.Not(customPropSetFilter);
				filterCollection.AddRange(new List<Exchange.SearchFilter> {
					lastSyncDateUtcFilter,
					notCustomPropSetFilter
				});
				return filterCollection;
			}
			return null;
		}

		public Exchange.Contact GetFullContact(Exchange.ItemId itemId) {
			return ExchangeUtility.SafeBindItem<Exchange.Contact>(Service, itemId);
		}

		public override ExchangeSettings GetSettings() {
			return UserSettings;
		}

		/// <summary>
		/// <see cref="RemoteProvider.NeedMetaDataActualization"/>
		/// </summary>
		public override bool NeedMetaDataActualization() {
			return !_userConnection.GetIsFeatureEnabled("ExchangeCalendarWithoutMetadata");
		}

		/// <summary>
		/// <see cref="RemoteProvider.CollectNewItems"/>
		/// </summary>
		public override IEnumerable<LocalItem> CollectNewItems(SyncContext context) {
			if (!UserSettings.ExportContacts) {
				return new List<LocalItem>();
			}
			SyncItemSchema primarySchema = SyncItemSchemaCollection.First(schema =>
					schema.PrimaryEntityConfig.Order == 0);
			return GetNotSyncedContacts(context, primarySchema);
		}

		/// <summary>
		/// Updates last synchronization date value.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		public override void CommitChanges(SyncContext context) {
			UpdateLastSyncDate(context.UserConnection, context);
			_syncErrorHelper.CleanUpSynchronizationError(SenderEmailAddress);
		}

		/// <summary>
		/// <see cref="RemoteProvider.CreateNewSyncItem"/>
		/// </summary>
		public override IRemoteItem CreateNewSyncItem(SyncItemSchema schema) {
			return new ExchangeContact(schema, new Exchange.Contact(Service), TimeZone) {
				Action = SyncAction.Create
			};
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.EnumerateChanges"/>
		/// </summary>
		public override IEnumerable<IRemoteItem> EnumerateChanges(SyncContext context) {
			base.EnumerateChanges(context);
			var result = new List<IRemoteItem>();
			if (!UserSettings.ImportContacts) {
				return result;
			}
			var folders = new List<Exchange.Folder>();
			if (UserSettings.ImportContactsAll) {
				Exchange.Folder rootFolder = Exchange.Folder.Bind(Service, Exchange.WellKnownFolderName.MsgFolderRoot);
				var filter = new Exchange.SearchFilter.IsEqualTo(
								Exchange.FolderSchema.FolderClass, ExchangeConsts.ContactFolderClassName);
				folders.GetAllFoldersByFilter(rootFolder, filter);
			} else {
				folders = SafeBindFolders(Service, UserSettings.RemoteFolderUIds.Keys, context);
			}
			Exchange.SearchFilter itemsFilter = GetExternalItemsFiltersHandler != null 
					? GetExternalItemsFiltersHandler() as Exchange.SearchFilter
					: GetContactsFilters();
			SyncItemSchema schema = FindSchemaBySyncValueName(typeof(ExchangeContact).Name);
			var exchangeAddrDetails = new List<ExchangeAddressDetail>();
			var exchangeCompanies = new List<string>();
			foreach (Exchange.Folder noteFolder in folders) {
				var itemView = new Exchange.ItemView(PageItemCount);
				Exchange.FindItemsResults<Exchange.Item> itemCollection;
				do {
					itemCollection = noteFolder.ReadItems(itemsFilter, itemView);
					foreach (Exchange.Item item in itemCollection) {
						Exchange.Contact fullContact = GetFullItemHandler != null 
								? GetFullItemHandler(item.Id.UniqueId) as Exchange.Contact
								: GetFullContact(item.Id);
						if (fullContact == null) {
							continue;
						}
						var remoteItem = new ExchangeContact(schema, fullContact, TimeZone);
						remoteItem.InitIdProperty(context);
						result.Add(remoteItem);
						if (!string.IsNullOrEmpty(fullContact.CompanyName)) {
							exchangeCompanies.Add(fullContact.CompanyName);
						}
						AddAddressDetails(exchangeAddrDetails, fullContact);
					}
				} while (itemCollection.MoreAvailable);
			}
			FillAddressDetailsData(context, exchangeAddrDetails);
			FillAccountsMap(context, exchangeCompanies.Distinct().ToList());
			return result;
		}

		/// <summary>
		/// <see cref="RemoteProvider.KnownTypes"/>
		/// </summary>
		public override IEnumerable<Type> KnownTypes() {
			return SyncTypeToTypeMap.Keys;
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.LoadSyncItem(SyncItemSchema, string)"/>
		/// </summary>
		public override IRemoteItem LoadSyncItem(SyncItemSchema schema, string id) {
			ExchangeBase remoteItem;
			string itemId = id.Split('_')[0];
			Exchange.Contact fullContact = ExchangeUtility.SafeBindItem<Exchange.Contact>(Service, new Exchange.ItemId(itemId));
			if (fullContact != null) {
				remoteItem = new ExchangeContact(schema, fullContact, TimeZone);
				remoteItem.Action = SyncAction.Update;
			} else {
				fullContact = new Exchange.Contact(Service);
				remoteItem = new ExchangeContact(schema, fullContact, id, TimeZone) {
					State = SyncState.Deleted
				};
			}
			return remoteItem;
		}

		/// <summary>
		/// <see cref="RemoteProvider.GetLocallyModifiedItemsMetadata"/>
		/// </summary>
		public override IEnumerable<ItemMetadata> GetLocallyModifiedItemsMetadata(SyncContext context,
				EntitySchemaQuery modifiedItemsEsq) {
			modifiedItemsEsq.Filters.Add(modifiedItemsEsq.CreateFilterWithParameters(FilterComparisonType.Equal,
					"CreatedBy", context.UserConnection.CurrentUser.ContactId));
			return base.GetLocallyModifiedItemsMetadata(context, modifiedItemsEsq);
		}

		/// <summary>
		/// ####### <see cref="Exchange.EmailAddressEntry"/> # ##### <see cref="Exchange.EmailAddressKey"/> 
		/// ## ####### ############# ######## #########.
		/// </summary>
		/// <param name="syncContext">######## #############.</param>
		/// <param name="contactRemoteId">########## ####### ####.</param>
		/// <param name="key">### ###### ######### #####, ## ######## ######## ######### 
		/// <see cref="Exchange.EmailAddressEntry"/>.</param>
		public void DeleteContactEmailAddress(SyncContext syncContext, string contactRemoteId,
				Exchange.EmailAddressKey key) {
			Exchange.ExtendedPropertyDefinition[] emailAddressProperties =
							_emailAddressPropertiesMap.GetExtendedPropertiesByKey(key);
			var propertyGroup = new List<Exchange.ExtendedPropertyDefinition>(emailAddressProperties);
			var propertySet = new Exchange.PropertySet(Exchange.BasePropertySet.IdOnly, propertyGroup);
			Exchange.Item contact;
			try {
				contact = Exchange.Item.Bind(Service, contactRemoteId, propertySet);
			} catch (Exception ex) {
				syncContext.LogError(SyncAction.Create, SyncDirection.Download,
					"[ExchangeContactSyncProviderImpl.DeleteContactEmailAddress]: Error loading contact with Id: {0}",
					ex, contactRemoteId);
				return;
			}
			foreach (Exchange.ExtendedPropertyDefinition prop in emailAddressProperties) {
				contact.RemoveExtendedProperty(prop);
			}
			try {
				contact.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite);
			} catch (Exception ex) {
				syncContext.LogError(SyncAction.Update, SyncDirection.Upload, string.Format(
					"[ExchangeContactSyncProviderImpl.DeleteContactEmailAddress] Error while updating contact with Id:" +
					"{0} Error details: {1}", contactRemoteId, ex.Message));
			}
		}

		/// <summary>
		/// <see cref="ExchangeSyncProvider.ApplyChanges"/>
		/// </summary>
		public override void ApplyChanges(SyncContext context, IRemoteItem syncItem) {
			base.ApplyChanges(context, syncItem);
			SyncAction action = syncItem.Action;
			if (action == SyncAction.Create) {
				((ExchangeContact)syncItem).InitIdProperty(context, true);
			}
		}

		/// <summary>
		/// Writes error message to the log, including information about the exception that caused this error.
		/// </summary>
		/// <param name="format">Format string with an informational message.</param>
		/// <param name="exception">Exeption.</param>
		/// <param name="args">Message format.</param>
		public override void LogError(string format, Exception exception, params object[] args) {
			base.LogError(format, exception, args);
			_syncErrorHelper.ProcessSynchronizationError(SenderEmailAddress, exception);
		}

		#endregion

	}

	#endregion

	#region Class: ExchangeContactAddressDetailHelper

	/// <summary>
	/// ############ ######### ##### ### ###### # ######## ######## # Exchange.
	/// </summary>
	internal class ExchangeContactAddressDetailHelper
	{

		#region Struct: RegionIdConfig

		private struct RegionIdConfig
		{
			public readonly Guid RegionId;

			public readonly Guid CountryId;

			public RegionIdConfig(Guid region, Guid country)
					: this() {
				RegionId = region;
				CountryId = country;
			}
		}

		#endregion

		#region Struct: RegionNameConfig

		private struct RegionNameConfig
		{
			public readonly string RegionName;

			public readonly string CountryName;

			public RegionNameConfig(string region, string country)
					: this() {
				RegionName = region;
				CountryName = country;
			}

		}

		#endregion

		#region Struct: CityIdConfig

		private struct CityIdConfig
		{
			public readonly Guid CityId;

			public RegionIdConfig RegionIdConfig;

			public CityIdConfig(Guid city, Guid region, Guid country)
					: this() {
				CityId = city;
				RegionIdConfig = new RegionIdConfig(region, country);
			}

		}

		#endregion

		#region Struct: CityNameConfig

		private struct CityNameConfig
		{
			public readonly string CityName;

			public RegionNameConfig RegionNameConfig;

			public CityNameConfig(string city, string region, string country)
					: this() {
				CityName = city;
				RegionNameConfig = new RegionNameConfig(region, country);
			}

		}

		#endregion

		#region Constants: Private

		private const string EmptyValueInKey = "#";

		#endregion

		#region Fields: Private

		private readonly UserConnection _userConnection;
		private readonly IEnumerable<ExchangeAddressDetail> _addressDetails;

		#endregion

		#region Constructors: Public

		/// <summary>
		/// ############## ##### ######### ########## ###### ### ###### # ######## ######## # Exchange.
		/// </summary>
		/// <param name="userConnection">######### ################# ###########.</param>
		/// <param name="addressDetails">###### ####### ###### ## Exchange.</param>
		public ExchangeContactAddressDetailHelper(UserConnection userConnection,
				IEnumerable<ExchangeAddressDetail> addressDetails) {
			_userConnection = userConnection;
			_addressDetails = addressDetails.Distinct();
		}

		#endregion

		#region Properties: Private

		private Dictionary<CityNameConfig, CityIdConfig> _cityMap;

		private Dictionary<CityNameConfig, CityIdConfig> CityMap {
			get {
				if (_cityMap != null) {
					return _cityMap;
				}
				IEnumerable<QueryParameter> cities = from city in _addressDetails
													 where !string.IsNullOrEmpty(city.CityName)
													 select new QueryParameter(city.CityName);
				_cityMap = GetCityMap(_userConnection, cities.ToList());
				return _cityMap;
			}
		}

		private Dictionary<string, Guid> _countryMap;

		private Dictionary<string, Guid> CountryMap {
			get {
				if (_countryMap != null) {
					return _countryMap;
				}
				IEnumerable<QueryParameter> countries = from country in _addressDetails
														where !string.IsNullOrEmpty(country.CountryName)
														select new QueryParameter(country.CountryName);
				_countryMap = GetCountryMap(_userConnection, countries.ToList());
				return _countryMap;
			}

		}

		private Dictionary<RegionNameConfig, RegionIdConfig> _regionMap;

		private Dictionary<RegionNameConfig, RegionIdConfig> RegionMap {
			get {
				if (_regionMap != null) {
					return _regionMap;
				}
				IEnumerable<QueryParameter> regions = from region in _addressDetails
													  where !string.IsNullOrEmpty(region.RegionName)
													  select new QueryParameter(region.RegionName);
				_regionMap = GetRegionMap(_userConnection, regions.ToList());
				return _regionMap;
			}

		}

		#endregion

		#region Methods: Private

		private static Dictionary<string, Guid> GetCountryMap(UserConnection userConnection,
				List<QueryParameter> countries) {
			var map = new Dictionary<string, Guid>();
			if (!countries.Any()) {
				return map;
			}
			var select = new Select(userConnection).Distinct()
				.Column("Country", "Name").As("CountryName")
				.Column("Country", "Id").As("CountryId")
				.From("Country")
				.Where("Country", "Name").In(countries.Distinct()) as Select;
			using (DBExecutor dbConnection = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbConnection)) {
					while (reader.Read()) {
						var countryName = reader.GetColumnValue<string>("CountryName");
						var countryId = reader.GetColumnValue<Guid>("CountryId");
						if (map.ContainsKey(countryName)) {
							continue;
						}
						map[countryName] = countryId;
					}
				}
			}
			return map;
		}

		private static Dictionary<RegionNameConfig, RegionIdConfig> GetRegionMap(UserConnection userConnection,
				List<QueryParameter> regions) {
			var map = new Dictionary<RegionNameConfig, RegionIdConfig>();
			if (!regions.Any()) {
				return map;
			}
			var select = new Select(userConnection).Distinct()
					.Column("Country", "Name").As("CountryName")
					.Column("Country", "Id").As("CountryId")
					.Column("Region", "Name").As("RegionName")
					.Column("Region", "Id").As("RegionId")
					.From("Region")
					.LeftOuterJoin("Country").On("Country", "Id").IsEqual("Region", "CountryId")
					.Where("Region", "Name").In(regions.Distinct()) as Select;
			using (DBExecutor dbConnection = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbConnection)) {
					while (reader.Read()) {
						var regionName = reader.GetColumnValue<string>("RegionName");
						var countryName = reader.GetColumnValue<string>("CountryName");
						var regionId = reader.GetColumnValue<Guid>("RegionId");
						var countryId = reader.GetColumnValue<Guid>("CountryId");
						var regionNameConfig = new RegionNameConfig(regionName, countryName);
						if (map.ContainsKey(regionNameConfig)) {
							continue;
						}
						map[regionNameConfig] = new RegionIdConfig(regionId, countryId);
					}
				}
			}
			return map;
		}

		private static Dictionary<CityNameConfig, CityIdConfig> GetCityMap(UserConnection userConnection,
				List<QueryParameter> cities) {
			var map = new Dictionary<CityNameConfig, CityIdConfig>();
			if (!cities.Any()) {
				return map;
			}
			var select = new Select(userConnection).Distinct()
					.Column("City", "Name").As("CityName")
					.Column("City", "Id").As("CityId")
					.Column("Country", "Name").As("CountryName")
					.Column("Country", "Id").As("CountryId")
					.Column("Region", "Name").As("RegionName")
					.Column("Region", "Id").As("RegionId")
					.From("City")
					.LeftOuterJoin("Country").On("City", "CountryId").IsEqual("Country", "Id")
					.LeftOuterJoin("Region").On("City", "RegionId").IsEqual("Region", "Id")
					.Where("City", "Name").In(cities.Distinct()) as Select;
			using (DBExecutor dbConnection = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbConnection)) {
					while (reader.Read()) {
						var cityName = reader.GetColumnValue<string>("CityName");
						var regionName = reader.GetColumnValue<string>("RegionName");
						var countryName = reader.GetColumnValue<string>("CountryName");
						var cityId = reader.GetColumnValue<Guid>("CityId");
						var regionId = reader.GetColumnValue<Guid>("RegionId");
						var countryId = reader.GetColumnValue<Guid>("CountryId");
						var cityNameConfig = new CityNameConfig(cityName, regionName, countryName);
						if (map.ContainsKey(cityNameConfig)) {
							continue;
						}
						map[cityNameConfig] = new CityIdConfig(cityId, regionId, countryId);
					}
				}
			}
			return map;
		}

		private void FillAddressDetail(AddressDetail addressDetail, CityIdConfig cityIdConfig) {
			addressDetail.CityId = cityIdConfig.CityId;
			addressDetail.RegionId = cityIdConfig.RegionIdConfig.RegionId;
			addressDetail.CountryId = cityIdConfig.RegionIdConfig.CountryId;
		}

		private void TryFillCity(ExchangeAddressDetail detail, AddressDetail addressDetail) {
			if (string.IsNullOrEmpty(detail.CityName)) {
				return;
			}
			var citiesValue = CityMap.Where(e => e.Key.CityName == detail.CityName).ToList();
			if (!citiesValue.Any()) {
				return;
			}
			if (citiesValue.Count() == 1) {
				FillAddressDetail(addressDetail, citiesValue.First().Value);
				return;
			}
			if (string.IsNullOrEmpty(detail.RegionName)) {
				return;
			}
			citiesValue = citiesValue.Where(e => e.Key.RegionNameConfig.RegionName == detail.RegionName).ToList();
			if (!citiesValue.Any()) {
				return;
			}
			if (citiesValue.Count() == 1) {
				FillAddressDetail(addressDetail, citiesValue.First().Value);
				return;
			}
			if (string.IsNullOrEmpty(detail.CountryName)) {
				return;
			}
			citiesValue = citiesValue.Where(e => e.Key.RegionNameConfig.CountryName == detail.CountryName).ToList();
			if (citiesValue.Count() == 1) {
				FillAddressDetail(addressDetail, citiesValue.First().Value);
			}
		}

		private void TryFillRegion(ExchangeAddressDetail detail, AddressDetail addressDetail) {
			if (string.IsNullOrEmpty(detail.RegionName) || addressDetail.RegionId != Guid.Empty) {
				return;
			}
			var regionsValue = RegionMap.Where(e => e.Key.RegionName == detail.RegionName).ToList();
			if (!regionsValue.Any()) {
				return;
			}
			KeyValuePair<RegionNameConfig, RegionIdConfig> regionValue;
			if (regionsValue.Count() == 1) {
				regionValue = regionsValue.First();
				addressDetail.RegionId = regionValue.Value.RegionId;
				if (addressDetail.CountryId == Guid.Empty) {
					addressDetail.CountryId = regionValue.Value.CountryId;
				}
			}
			if (string.IsNullOrEmpty(detail.CountryName)) {
				return;
			}
			regionsValue = regionsValue.Where(e => e.Key.CountryName == detail.CountryName).ToList();
			if (!regionsValue.Any() || regionsValue.Count() != 1) {
				return;
			}
			regionValue = regionsValue.First();
			addressDetail.RegionId = regionValue.Value.RegionId;
			if (addressDetail.CountryId == Guid.Empty) {
				addressDetail.CountryId = regionValue.Value.CountryId;
			}
		}

		private void TryFillCountry(ExchangeAddressDetail detail, AddressDetail addressDetail) {
			if (string.IsNullOrEmpty(detail.CountryName) || addressDetail.CountryId != Guid.Empty) {
				return;
			}
			if (CountryMap.ContainsKey(detail.CountryName)) {
				addressDetail.CountryId = CountryMap[detail.CountryName];
			}
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// ########## ########## #### ## ########### ###### ######## # Exchange
		/// <see cref="Exchange.PhysicalAddressEntry"/>.
		/// </summary>
		/// <param name="physicalAddressEntry">##### ######## # Exchange.</param>
		public static string GetUniqueKey(Exchange.PhysicalAddressEntry physicalAddressEntry) {
			var addressDetail = new ExchangeAddressDetail(physicalAddressEntry.CountryOrRegion,
				physicalAddressEntry.State, physicalAddressEntry.City);
			return GetUniqueKey(addressDetail);
		}

		/// <summary>
		/// ########## ########## ####, ############ ## ######## ######, ######## ####### # ######## ######,
		/// ### ########## ###### ###### ######## Exchange <see cref="ExchangeAddressDetail"/>.
		/// </summary>
		/// <param name="addressDetail">###### ###### ######## Exchange.</param>
		public static string GetUniqueKey(ExchangeAddressDetail addressDetail) {
			var key = new StringBuilder();
			key.Append(string.IsNullOrEmpty(addressDetail.CountryName) ? EmptyValueInKey : addressDetail.CountryName);
			key.Append(string.IsNullOrEmpty(addressDetail.RegionName) ? EmptyValueInKey : addressDetail.RegionName);
			key.Append(string.IsNullOrEmpty(addressDetail.CityName) ? EmptyValueInKey : addressDetail.CityName);
			return key.ToString();
		}

		/// <summary>
		/// ########## ######## ###### ###### ######## ########## ######### <see cref="AddressDetail"/> ## #####,
		/// ############# ## ######## ######, ######## ####### # ######## ######.
		/// </summary>
		public Dictionary<string, AddressDetail> GetAddressesLookupMap() {
			var map = new Dictionary<string, AddressDetail>();
			foreach (ExchangeAddressDetail detail in _addressDetails) {
				string key = GetUniqueKey(detail);
				if (map.ContainsKey(key)) {
					continue;
				}
				var addressDetail = new AddressDetail(Guid.Empty, Guid.Empty, Guid.Empty);
				TryFillCity(detail, addressDetail);
				TryFillRegion(detail, addressDetail);
				TryFillCountry(detail, addressDetail);
				map[key] = addressDetail;
			}
			return map;
		}

		#endregion

	}

	#endregion

	#region Struct: ExchangeAddressDetail

	/// <summary>
	/// ###### ###### ######## Exchange.
	/// </summary>
	internal struct ExchangeAddressDetail
	{

		#region Constructors: Public

		/// <summary>
		/// ############## ##### ######### ####### <see cref="ExchangeAddressDetail"/> ######### ###########.
		/// </summary>
		/// <param name="cityName">######## ######.</param>
		/// <param name="countryName">######## ######.</param>
		/// <param name="regionName">######## #######.</param>
		public ExchangeAddressDetail(string countryName, string regionName, string cityName)
			: this() {
			CityName = cityName;
			CountryName = countryName;
			RegionName = regionName;
		}

		#endregion

		#region Properties: Public

		/// <summary>
		/// ######## ######.
		/// </summary>
		public string CityName {
			get;
			private set;
		}

		/// <summary>
		/// ######## ######.
		/// </summary>
		public string CountryName {
			get;
			private set;
		}

		/// <summary>
		/// ######## #######.
		/// </summary>
		public string RegionName {
			get;
			private set;
		}

		#endregion

	}

	#endregion


}