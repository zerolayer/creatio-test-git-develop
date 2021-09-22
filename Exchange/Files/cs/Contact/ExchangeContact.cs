namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Linq;
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Common;
	using Terrasoft.Sync;
	using Terrasoft.Configuration;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeContact

	/// <summary>
	/// Provides methods for Contact synchronization with Exchange.
	/// </summary>
	[Map("Contact", 0, IsPrimarySchema = true, Direction = SyncDirection.DownloadAndUpload,
			FetchColumnNames = new[] { "Name", "Surname", "GivenName", "MiddleName", "Account", "JobTitle",
			"Department", "BirthDate", "Owner", "SalutationType", "Gender" })]
	[Map("ContactCommunication", 1, PrimarySchemaName = "Contact", ForeingKeyColumnName = "Contact",
			Direction = SyncDirection.DownloadAndUpload, FetchColumnNames
				= new[] { "Number", "CreatedOn", "CommunicationType" })]
	[Map("ContactAddress", 1, PrimarySchemaName = "Contact", ForeingKeyColumnName = "Contact",
			Direction = SyncDirection.DownloadAndUpload, FetchColumnNames
				= new[] { "City", "Country", "Region", "Address", "Zip", "AddressType", "CreatedOn" })]
	public class ExchangeContact : ExchangeBase
	{

		#region Fields: Private

		private static Exchange.PropertySet _propertySet =
			new Exchange.PropertySet(Exchange.BasePropertySet.FirstClassProperties);

		#endregion
		
		#region Constructors: Public

		/// <summary>
		/// Creates new <see cref="ExchangeContact"/> instance, using <paramref name="schema"/> and remote storage item.
		/// </summary>
		/// <param name="schema">Entity sync schema instance.</param>
		/// <param name="item">Remote storage item.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeContact(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timeZoneInfo)
			: base(schema, item, timeZoneInfo) {
			_propertySet.RequestedBodyType = Exchange.BodyType.Text;
			_propertySet.Add(ExchangeUtilityImpl.LocalIdProperty);
		}

		/// <summary>
		/// Creates new  <see cref="ExchangeContact"/> instance, using <paramref name="schema"/>, remote storage item 
		/// instance and remete storage item id.
		/// </summary>
		/// <param name="schema">Entity sync schema instance.</param>
		/// <param name="item">Remote storage item.</param>
		/// <param name="remoteId">Remote storage item id.</param>
		/// <param name="timeZoneInfo">Current user timezone.</param>
		public ExchangeContact(SyncItemSchema schema, Exchange.Item item, string remoteId, TimeZoneInfo timeZoneInfo)
			: this(schema, item, timeZoneInfo) {
			RemoteId = remoteId;
		}

		#endregion

		#region Constants: Private

		private const int TitleTagId = 0x3A45;
		private const int GenderTagId = 0x3A4D;
		private const int ExtendedPropertyGenderMaleValue = 2;
		private const int ExtendedPropertyGenderFemaleValue = 1;
		private const int ExtendedPropertyGenderNotSetValue = 3;

		#endregion

		#region Properties: Public

		/// <summary>
		/// Unique appointment item id in remote storage.
		/// </summary>
		private string _id;
		public override string Id {
			get {
				if (RemoteId.IsNullOrEmpty()) {
					RemoteId = Item.Id.UniqueId;
				}
				return _id.IsNullOrEmpty() ? RemoteId : _id;
			}
			internal set {
				_id = value;
			}
		}
		
		#endregion

		#region Methods: Private

		/// <summary>
		/// Methods returns unique id for <c>Web</c> <see cref="ContactCommunication"/> record, returns first by value
		/// or CreatedOn date. Returns <c>Guid.Empty</c> if records not found.
		/// </summary>
		/// <param name="userConnection">User connection instance.</param>
		/// <param name="contactId"><see cref="Contact"/> unique id.</param>
		/// <param name="value">Communication search value. Default value is <c>null</c>.</param>
		/// <returns>
		/// Unique id for <c>Web</c> <see cref="ContactCommunication"/> record or <c>Guid.Empty</c> if
		/// records not found.
		///</returns>
		private static Guid TryFindWebPageDetailInLocalStore(UserConnection userConnection,
				Guid contactId, string value = null) {
			var id = Guid.Empty;
			Select select = new Select(userConnection).Top(1)
							.Column("Id")
							.From("ContactCommunication")
							.Where("ContactId").IsEqual(Column.Parameter(contactId))
							.And("CommunicationTypeId").IsEqual(Column.Parameter(Guid.Parse(CommunicationTypeConsts.WebId)))
							.OrderBy(OrderDirectionStrict.Ascending, "CreatedOn") as Select;
			if (!string.IsNullOrEmpty(value)) {
				select =  select.And("SearchNumber").IsEqual(Column.Parameter(value.ToLower())) as Select;
			}
			using (DBExecutor dbExecutor = userConnection.EnsureDBConnection()) {
				using (IDataReader reader = select.ExecuteReader(dbExecutor)) {
					if (reader.Read()) {
						id = reader.GetColumnValue<Guid>("Id");
					}
				}
			}
			return id;
		}

		/// <summary>
		/// Creates o replaces <c>Web</c> <see cref="ContactCommunication"/> instance, using <paramref name="value"/>, 
		/// in <paramref name="localItem"/> sync entities collection.
		/// </summary>
		/// <param name="userConnection">User connection instance.</param>
		/// <param name="contactId"><see cref="Contact"/> unique id.</param>
		/// <param name="value">ContactCommunication value.</param>
		/// <param name="localItem">Local storage synchronization element.</param>
		private static void AddOrReplaceWebPageInstance(UserConnection userConnection, string value,
				Guid contactId, LocalItem localItem) {
			EntitySchema schema = userConnection.EntitySchemaManager.GetInstanceByName("ContactCommunication");
			var instance = (ContactCommunication)schema.CreateEntity(userConnection);
			Guid id = TryFindWebPageDetailInLocalStore(userConnection, contactId, value);
			if (id == Guid.Empty) {
				instance.SetDefColumnValues();
			} else if (!instance.FetchFromDB(id, false)) {
				instance.SetDefColumnValues();
			}
			instance.SetColumnValue("Number", value);
			instance.SetColumnValue("CommunicationTypeId", CommunicationTypeConsts.WebId);
			instance.SetColumnValue("ContactId", contactId);
			localItem.AddOrReplace(schema.Name, SyncEntity.CreateNew(instance));
		}

		private void AddOrReplaceWebPageInLocalItem(UserConnection userConnection, LocalItem localItem,
				Guid mainRecordId, string webPageValue) {
			IEnumerable<SyncEntity> webPageEntities = localItem.Entities["ContactCommunication"].Where(e =>
							e.State != SyncState.New && e.State != SyncState.Deleted &&
							e.Entity.GetTypedColumnValue<Guid>("CommunicationTypeId") ==
								Guid.Parse(CommunicationTypeConsts.WebId));
			if (webPageEntities.Any()) {
				SyncEntity webPageSyncEntity = webPageEntities.First();
				if (string.IsNullOrEmpty(webPageValue)) {
					webPageSyncEntity.Action = SyncAction.Delete;
				} else {
					webPageSyncEntity.Entity.SetColumnValue("Number", webPageValue);
					webPageSyncEntity.Action = SyncAction.Update;
				}
			} else if (!string.IsNullOrEmpty(webPageValue)) {
				AddOrReplaceWebPageInstance(userConnection, webPageValue, mainRecordId, localItem);
			}
		}

		private string GetWebPageForRemoteItem(LocalItem localItem, UserConnection userConnection, Guid contactId) {
			string result = string.Empty;
			IEnumerable<SyncEntity> webPageEntities = localItem.Entities["ContactCommunication"].Where(e =>
							e.State != SyncState.New && e.State != SyncState.Deleted &&
							e.Entity.GetTypedColumnValue<Guid>("CommunicationTypeId") ==
								Guid.Parse(CommunicationTypeConsts.WebId));
			if (webPageEntities.Any()) {
				SyncEntity webPageSyncEntity = webPageEntities.First();
				result = webPageSyncEntity.Entity.GetTypedColumnValue<string>("Number");
			} else {
				Guid id = TryFindWebPageDetailInLocalStore(userConnection, contactId);
				if (id == Guid.Empty) {
					return result;
				}
				EntitySchema schema = userConnection.EntitySchemaManager.GetInstanceByName("ContactCommunication");
				var instance = (ContactCommunication)schema.CreateEntity(userConnection);
				if (instance.FetchFromDB(id, false)) {
					localItem.AddOrReplace(schema.Name, SyncEntity.CreateNew(instance));
					result = instance.Number;
				}
			}
			return result;
		}

		private void SetContactGender(Exchange.Contact contact, Guid genderId) {
			if (genderId == Guid.Empty) {
				return;
			}
			var extendedPropertyDefinition = new Exchange.ExtendedPropertyDefinition(GenderTagId,
				Exchange.MapiPropertyType.Short);
			switch (genderId.ToString().ToUpper()) {
				case ContactConsts.MaleId:
					contact.SetExtendedProperty(extendedPropertyDefinition, ExtendedPropertyGenderMaleValue);
					break;
				case ContactConsts.FemaleId:
					contact.SetExtendedProperty(extendedPropertyDefinition, ExtendedPropertyGenderFemaleValue);
					break;
				default:
					contact.SetExtendedProperty(extendedPropertyDefinition, ExtendedPropertyGenderNotSetValue);
					break;
			}
		}

		private static void SetContactTitle(Exchange.Contact contact, string title) {
			if (string.IsNullOrEmpty(title)) {
				return;
			}
			var titleExtendedProperty = new Exchange.ExtendedPropertyDefinition(TitleTagId,
				Exchange.MapiPropertyType.String);
			contact.SetExtendedProperty(titleExtendedProperty, title);
		}

		private List<SyncEntity> GetContactSyncItems(SyncContext context, LocalItem localItem, Guid storedId) {
			var contacts = localItem.Entities["Contact"];
			if (storedId.IsEmpty() || contacts.Count > 0) {
				return contacts;
			}
			EntitySchema schema = context.UserConnection.EntitySchemaManager.GetInstanceByName("Contact");
			Entity existingContact = schema.CreateEntity(context.UserConnection);
			if (existingContact.FetchFromDB(storedId)) {
				localItem.AddOrReplace("Contact", new SyncEntity(existingContact, SyncState.None) {
					Action = SyncAction.Update
				});
			}
			return contacts;
			
		}

		#endregion

		#region Methods: Protected

		protected override string GetDisplayName() {
			var contactExchange = Item as Exchange.Contact;
			if (contactExchange == null) {
				return Id;
			}
			var displayName = contactExchange.SafeGetValue<string>(Exchange.ContactSchema.DisplayName);
			return !string.IsNullOrEmpty(displayName) ? displayName : Id;
		}
		
		/// <summary>
		/// Check value of name fields of <paramref name="exchangeContact"/>.
		/// </summary>
		/// <param name="exchangeContact"><see cref="Exchange.Contact"/> instance.</param>
		/// <param name="context">Sync context.</param>
		/// <returns>
		/// True if all name fields empty.
		/// </returns>
		protected bool GetIsContactNameEmpty(Exchange.Contact exchangeContact, SyncContext context) {
			string surname = exchangeContact.Surname;
			string givenName = exchangeContact.GivenName;
			string middleName = exchangeContact.MiddleName;
			bool result = (string.IsNullOrEmpty(surname) && string.IsNullOrEmpty(givenName) &&
					string.IsNullOrEmpty(middleName));
			if (result) {
				context.LogError(Action, SyncDirection.Upload,
					GetLocalizableString(context.UserConnection, "NameFieldsEmptyTpl"), GetItemSimpleId());
			}
			return result;
		}

		/// <summary>
		/// Checks if contact metadata exists by <paramref name="contactId"/>
		/// </summary>
		/// <param name="contactId">Contact uniqueidentifier.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>True if contact metadata exists. </returns>
		protected virtual bool IsContactMetadataExist(Guid contactId, UserConnection userConnection) {
			return ExchangeUtility.IsContactMetadataExist(contactId, userConnection);
		}

		/// <summary>
		/// Returns <see cref="Configuration.Contact"/> instance, which represents <paramref name="exchangeContact"/>.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem">Local storage synchronization element.</param>
		/// <param name="exchangeContact"><see cref="Exchange.Contact"/> instance.</param>
		/// <returns><see cref="Configuration.Contact"/> instance.</returns>
		protected Entity GetContactInstance(SyncContext context, LocalItem localItem,
				Exchange.Contact exchangeContact) {
			Guid storedId;
			var hasStoredId = GetContactLocalIdProperty(out storedId);
			List<SyncEntity> contacts = GetContactSyncItems(context, localItem, storedId);
			var localContactId = contacts.Count > 0 ? contacts[0].EntityId : Guid.Empty;
			bool isDublicateByRemoteId = ((!hasStoredId || storedId != localContactId) && contacts.Count > 0);
			if (isDublicateByRemoteId && IsContactMetadataExist(storedId, context.UserConnection)) {
				return null;
			}
			if (isDublicateByRemoteId) {
				foreach(KeyValuePair<string, List<SyncEntity>> entities in localItem.Entities) {
					entities.Value.Clear();
				}
			}
			Entity instance = GetEntityInstance<Entity>(context, localItem, "Contact");
			if (isDublicateByRemoteId) {
				Id = GetExtendedIdString(instance.PrimaryColumnValue);
			}
			return instance;
		}

		/// <summary>
		/// Generates unique contact id string.
		/// </summary>
		/// <param name="contactId"><see cref="Configuration.Contact"/> instance id.</param>
		/// <returns>Unique contact id string.</returns>
		protected virtual string GetExtendedIdString(Guid contactId) {
			return GetItemSimpleId() + "_" + contactId.ToString().ToUpper();
		}

		/// <summary>
		/// Returns true if <see cref="ExchangeBase.Item"/> instance contains <see cref="ExchangeUtilityImpl.LocalIdProperty"/>
		/// value. Sets the resulting value in <paramref name="localIdProperty"/> parameter.
		/// </summary>
		/// <param name="localIdProperty"><see cref="ExchangeUtilityImpl.LocalIdProperty"/> property value.</param>
		/// <returns>
		/// True if <see cref="ExchangeBase.Item"/> instance contains <see cref="ExchangeUtilityImpl.LocalIdProperty"/>,
		/// false otherwise.
		/// <returns>
		/// <remarks>
		/// External dependency.
		/// </remarks>
		protected virtual bool GetContactLocalIdProperty(out Guid localIdProperty) {
			Object localId;
			if (!Item.TryGetProperty(ExchangeUtilityImpl.LocalIdProperty, out localId)) {
				localIdProperty = Guid.Empty;
				return false;
			}
			localIdProperty = Guid.Parse(localId.ToString());
			return true;
		}
		
		/// <summary>
		/// Calls <see cref="Exchange.Item.Load(Exchange.PropertySet)"/> method for <paramref name="exchangeContact"/>.
		/// </summary>
		/// <param name="exchangeContact"><see cref="Exchange.Contact"/> instance.</param>
		/// <remarks>
		/// External dependency allocation.
		/// </remarks>
		protected virtual void LoadItemProperties(Exchange.Contact exchangeContact) {
			exchangeContact.Load(_propertySet);
		}

		/// <summary>
		/// Checks if local contact instance deleted.
		/// </summary>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <returns>
		/// Returns true if contact entity deleted, false otherwise.
		/// </returns>
		protected bool IsLocalItemDeleted(ref LocalItem localItem) {
			List<SyncEntity> contacts = localItem.Entities["Contact"];
			if (contacts.Count > 0 && contacts[0].State == SyncState.Deleted) {
				return true;
			}
			return false;
		}
		
		/// <summary>
		/// Returns localizable string value.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="stringName">Localizable string name.</param>
		/// <returns>
		/// Localizable string value.
		/// </returns>
		protected string GetLocalizableString(UserConnection userConnection, string stringName) {
			return new LocalizableString(userConnection.ResourceStorage, "ExchangeContact",
					string.Format("LocalizableStrings.{0}.Value", stringName)).ToString();
		}
		
		/// <summary>
		/// Returns exchange item unique id.
		/// </summary>
		/// <returns>
		/// Exchange item unique id.
		/// </returns>
		protected virtual string GetItemSimpleId() {
			return Item.Id.UniqueId;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Fills element synchronization in the local storage <paramref name="localItem"/> 
		/// the value of the element in the external storage.
		/// </summary>
		/// <param name="localItem">Local storage synchronization element.</param>
		/// <param name="context">Synchronization context.</param>
		public override void FillLocalItem(SyncContext context, ref LocalItem localItem) {
			if (IsDeletedProcessed("Contact", ref localItem) || IsLocalItemDeleted(ref localItem)) {
				return;
			}
			var contactRemoteProvider = context.RemoteProvider as ExchangeContactSyncProviderImpl;
			var exchangeContact = Item as Exchange.Contact;
			if (exchangeContact == null) {
				context.LogError(Action, SyncDirection.Upload,
					GetLocalizableString(context.UserConnection, "InvalidItemTypeTpl"), GetItemSimpleId(),
					Item.Subject);
				Action = SyncAction.None;
				return;
			}
			string surname = exchangeContact.Surname;
			string givenName = exchangeContact.GivenName;
			string middleName = exchangeContact.MiddleName;
			if (GetIsContactNameEmpty(exchangeContact, context) || GetRemoteItemLockedForSync(context)) {
				Action = SyncAction.None;
				return;
			}
			LoadItemProperties(exchangeContact);
			var contact = GetContactInstance(context, localItem, exchangeContact);
			if (contact == null || !SetContactExtendProperty(context, exchangeContact, contact.PrimaryColumnValue)) {
				Action = SyncAction.None;
				return;
			}
			contact.SetColumnValue("Surname", surname);
			contact.SetColumnValue("GivenName", givenName);
			contact.SetColumnValue("MiddleName", middleName);
			contact.SetColumnValue("OwnerId", context.UserConnection.CurrentUser.ContactId);
			contact.SetColumnValue("JobTitle", exchangeContact.JobTitle);
			var birthdate = exchangeContact.SafeGetValue<DateTime>(Exchange.ContactSchema.Birthday);
			if (birthdate != DateTime.MinValue) {
				contact.SetColumnValue("BirthDate", birthdate);
			}
			Dictionary<string, Guid> accountsMap = contactRemoteProvider.AccountsMap;
			if (accountsMap != null) {
				string companyName = exchangeContact.CompanyName;
				if (string.IsNullOrEmpty(companyName)) {
					contact.SetColumnValue("AccountId", null);
				} else if (accountsMap.ContainsKey(companyName)) {
					contact.SetColumnValue("AccountId", accountsMap[companyName]);
				}
			}
			AddOrReplaceWebPageInLocalItem(context.UserConnection, localItem, contact.PrimaryColumnValue,
				exchangeContact.BusinessHomePage);
			var emailAddressesSynchronizer = new ExchangeEmailAddressDetailsSynchronizer(context, localItem,
				exchangeContact);
			var addressesSynchronizer = new ExchangeAddressDetailsSynchronizer(context, localItem, exchangeContact);
			var phoneNumbersSynchronizer = new ExchangePhoneNumbersDetailsSynchronizer(context, localItem,
				exchangeContact);
			emailAddressesSynchronizer.SyncLocalDetails();
			addressesSynchronizer.SyncLocalDetails();
			phoneNumbersSynchronizer.SyncLocalDetails();
		}

		/// <summary>
		/// Fills element in the remote storage from the element synchronization
		/// in the local storage.<paramref name="localItem"/>.
		/// </summary>
		/// <param name="localItem">The element synchronization in the local storage.</param>
		/// <param name="context">Synchronization context.</param>
		public override void FillRemoteItem(SyncContext context, LocalItem localItem) {
			var exchangeContact = (Exchange.Contact)Item;
			SyncEntity localEntity = localItem.Entities["Contact"][0];
			if (localEntity == null) {
				return;
			}
			if (localEntity.State == SyncState.Deleted) {
				return;
			}
			if (Action == SyncAction.None) {
				return;
			}
			var localContact = localEntity.Entity;
			if (localContact == null || GetEntityLockedForSync(localContact.PrimaryColumnValue, context)) {
				Action = SyncAction.None;
				return;
			}
			if (Action == SyncAction.Update) {
				InitIdProperty(context);
			}
			exchangeContact.DisplayName = localContact.GetTypedColumnValue<string>("Name");
			exchangeContact.Surname = localContact.GetTypedColumnValue<string>("Surname");
			exchangeContact.GivenName = localContact.GetTypedColumnValue<string>("GivenName");
			exchangeContact.MiddleName = localContact.GetTypedColumnValue<string>("MiddleName");
			exchangeContact.CompanyName = localContact.GetTypedColumnValue<string>("AccountName");
			exchangeContact.JobTitle = localContact.GetTypedColumnValue<string>("JobTitle");
			exchangeContact.Department = localContact.GetTypedColumnValue<string>("DepartmentName");
			var birthDate = localContact.GetTypedColumnValue<DateTime>("BirthDate");
			if (birthDate != DateTime.MinValue) {
				exchangeContact.Birthday = birthDate;
			}
			if (Action == SyncAction.Create) {
				exchangeContact.FileAsMapping = Exchange.FileAsMapping.DisplayName;
				exchangeContact.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, localContact.PrimaryColumnValue.ToString());
			}
			exchangeContact.BusinessHomePage = GetWebPageForRemoteItem(localItem, context.UserConnection,
				localContact.PrimaryColumnValue);
			SetContactTitle(exchangeContact, localContact.GetTypedColumnValue<string>("SalutationTypeName"));
			SetContactGender(exchangeContact, localContact.GetTypedColumnValue<Guid>("GenderId"));
			var emailAddressesSynchronizer =
					new ExchangeEmailAddressDetailsSynchronizer(context, localItem, exchangeContact);
			var addressesSynchronizer =
					new ExchangeAddressDetailsSynchronizer(context, localItem, exchangeContact);
			var phoneNumbersSynchronizer =
					new ExchangePhoneNumbersDetailsSynchronizer(context, localItem, exchangeContact);
			emailAddressesSynchronizer.SyncRemoteDetails();
			addressesSynchronizer.SyncRemoteDetails();
			phoneNumbersSynchronizer.SyncRemoteDetails();
		}

		/// <summary>
		/// Inititalezes unique exchange contact id property.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="useSimpleId">If true, uses <see cref="ExchangeContact.GetItemSimpleId"/> method result 
		/// as search value.</param>
		public void InitIdProperty(SyncContext context, bool useSimpleId = false) {
			string extendedId = GetExtendedId();
			string itemId = GetItemSimpleId();
			string searchId = useSimpleId ? itemId : extendedId;
			Select select = new Select(context.UserConnection)
				.Column(Func.Count("Id"))
				.From("SysSyncMetaData")
				.Where("RemoteId").IsEqual(Column.Parameter(searchId)) as Select;
			int count = select.ExecuteScalar<int>();
			Id = count > 0 ? extendedId : itemId;
		}

		/// <summary>
		/// Returrns current instance extended id.
		/// </summary>
		/// <returns>
		/// Unique exchange contact id.
		/// </returns>
		public string GetExtendedId() {
			Guid localId;
			LoadItemProperties(Item as Exchange.Contact);
			if (!GetContactLocalIdProperty(out localId)) {
				return GetItemSimpleId();
			}
			return GetExtendedIdString(localId);
		}

		/// <summary>
		/// Sets contact extend property.
		/// </summary>
		/// <param name="exchangeContact">Exchange contact item in external storage.</param>
		/// <param name="contact">Contact item.</param>
		/// <param name="context">Synchronization context.</param>
		/// <returns>Status flag setting an extend property.</returns>
		public virtual bool SetContactExtendProperty(SyncContext context, Exchange.Contact exchangeContact, Guid contactId) {
			try {
				Object localId;
				if (!exchangeContact.TryGetProperty(ExchangeUtilityImpl.LocalIdProperty, out localId)) {
					exchangeContact.SetExtendedProperty(ExchangeUtilityImpl.LocalIdProperty, contactId.ToString());
					exchangeContact.Update(Exchange.ConflictResolutionMode.AlwaysOverwrite);
				}
				return true;
			} catch (Exception ex) {
				context.LogError(Action, SyncDirection.Upload,
					GetLocalizableString(context.UserConnection, "SetExtendedPropertyErrorTpl"), ex.Message);
				return false;
			}
		}

		#endregion

	}

	#endregion

}