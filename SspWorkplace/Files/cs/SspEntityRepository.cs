namespace Terrasoft.Configuration.SspEntity
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Configuration.Section;
	using Terrasoft.Nui.ServiceModel.WebService;

	#region Class: SspEntityRepository

	[DefaultBinding(typeof(ISspEntityRepository))]
	public class SspEntityRepository : ISspEntityRepository {

		#region Fields: Private

		/// <summary>
		/// <see cref="SchemaDesignerUtilities"/> instance.
		/// </summary>
		private readonly SchemaDesignerUtilities _utils;

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		/// <summary>
		/// <see cref="EntitySchemaManager"/> instance.
		/// </summary>
		private readonly EntitySchemaManager _entitySchemaManager;

		/// <summary>
		/// <see cref="DBSecurityEngine"/> instance.
		/// </summary>
		private readonly DBSecurityEngine _securityEngine;

		/// <summary>
		/// <see cref="IResourceStorage"/> implementation instance.
		/// </summary>
		private readonly IResourceStorage _resourceStorage;

		/// <summary>
		/// Current <see cref="SysPackage"/> identifier.
		/// </summary>
		private readonly Guid _packageUId;

		#endregion

		#region Constructors: Public

		public SspEntityRepository(UserConnection uc) {
			_userConnection = uc;
			_entitySchemaManager = _userConnection.EntitySchemaManager;
			_securityEngine = _userConnection.DBSecurityEngine;
			_resourceStorage = uc.ResourceStorage;
			_utils = new SchemaDesignerUtilities(_userConnection);
			_packageUId = GetCurrentPackageUId();
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Returns avaliable for SSP users columns from <paramref name="entityUId"/>.
		/// </summary>
		/// <param name="entityUId"><see cref="Entity"/> unique identifier.</param>
		/// <returns>Columns unique identifiers list.</returns>
		private IEnumerable<Guid> GetSspColumnAccessListFromDb(Guid entityUId) {
			var select = new Select(_userConnection)
				.Column("ColumnUId")
				.From("PortalColumnAccessList").As("pcal")
				.InnerJoin("PortalSchemaAccessList").As("psal").On("pcal", "PortalSchemaListId").IsEqual("psal", "Id")
				.Where("psal", "SchemaUId").IsEqual(Column.Parameter(entityUId)) as Select;
			List<Guid> result = new List<Guid>();
			using (DBExecutor dbExecutor = _userConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						result.AddIfNotExists(dataReader.GetColumnValue<Guid>("ColumnUId"));
					}
				}
			}
			return result;
		}

		/// <summary>
		/// Returns current package unique identifier.
		/// </summary>
		/// <returns><see cref="SysPackage"/> unique identifier.</returns>
		private Guid GetCurrentPackageUId() {
			return Terrasoft.Core.Configuration.SysSettings.GetValue(_userConnection, "CurrentPackageId", Guid.Empty);
		}

		/// <summary>
		/// Sets <see cref="EntitySchema"/> property <paramref name="propertyName"/> enabled.
		/// </summary>
		/// <param name="item"><see cref="ISchemaManagerItem"/> implementation instance.</param>
		/// <param name="propertyName"><see cref="EntitySchema"/> property name.</param>
		private void SetEntitySchemaPropertyEnabled(ISchemaManagerItem<EntitySchema> item, string propertyName) {
			if (item == null) {
				return;
			}
			var designItem = GetDesignItem(item);
			if (!designItem.Instance.HasProperty(propertyName)) {
				return;
			}
			var propertyValue = designItem.Instance.GetPropertyValue(propertyName) as bool?;
			if (!propertyValue.HasValue || propertyValue.Value) {
				return;
			}
			designItem.Instance.SetPropertyValue(propertyName, true);
			_entitySchemaManager.SaveSchema(designItem, _userConnection);
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Gets <see cref="ISchemaManagerItem"/> manager design item for <paramref name="item"/>.
		/// </summary>
		/// <param name="item"><see cref="EntitySchemaManager"/> manager item.</param>
		/// <returns><see cref="ISchemaManagerItem"/> manager design item.</returns>
		protected virtual ISchemaManagerItem<EntitySchema> GetDesignItem(ISchemaManagerItem<EntitySchema> item) {
			ISchemaManagerItem managerItem;
			try {
				(Guid schemaId, Guid schemaRealUId) = _utils.GetSchemaPackageInfo(item.Name, _packageUId);
				managerItem = _entitySchemaManager.DesignItem(_userConnection, schemaRealUId);
				managerItem.Id = schemaId;
				var instance = managerItem.Instance as EntitySchema;
				instance.Id = schemaId;
			} catch (NullReferenceException) {
				managerItem = _entitySchemaManager.CreateDesignSchema(_userConnection,
					item.UId, _packageUId, true);
			}
			return _entitySchemaManager.FindDesignItem(_userConnection, managerItem.UId) as ISchemaManagerItem<EntitySchema>;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="ISspEntityRepository.GetSspColumnAccessList(Guid)"/>
		/// </summary>
		public IEnumerable<Guid> GetSspColumnAccessList(Guid entityUId) {
			return GetSspColumnAccessListFromDb(entityUId);
		}

		/// <summary>
		/// <see cref="ISspEntityRepository.IsEntitySspAllowed(string)"/>
		/// </summary>
		public virtual bool IsEntitySspAllowed(string entityName) {
			return !_securityEngine.GetIsLicensedEntity(entityName);
		}

		/// <summary>
		/// <see cref="ISspEntityRepository.SetEntitySspAllowed(Guid)"/>
		/// </summary>
		public void SetEntitySspAllowed(Guid entityUId) {
			var item = _entitySchemaManager.FindItemByUId(entityUId);
			if (IsEntitySspAllowed(item.Name)) {
				SetEntitySchemaPropertyEnabled(item, "IsSSPAvailable");
			}
		}

		/// <summary>
		/// <see cref="ISspEntityRepository.SetEntityAdministratedByOperations(Guid)"/>
		/// </summary>
		public void SetEntityAdministratedByOperations(Guid entityUId) {
			var item = _entitySchemaManager.FindItemByUId(entityUId);
			SetEntitySchemaPropertyEnabled(item, "AdministratedByOperations");
			_securityEngine.SetEntitySchemaOperationsRightLevel(UsersConsts.AllEmployersSysAdminUnitUId, item.UId, SchemaOperationRightLevels.All);
			_securityEngine.SetEntitySchemaOperationsRightLevel(UsersConsts.PortalUsersSysAdminUnitUId, item.UId, SchemaOperationRightLevels.All);
			_securityEngine.SetEntitySchemaAdministratedByOperations(item.UId, true);
			_securityEngine.SetEntitySchemaAdministratedByOperations(item.Name, true);
		}

		#endregion

	}

	#endregion

}
 