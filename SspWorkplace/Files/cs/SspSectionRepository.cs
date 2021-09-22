namespace Terrasoft.Configuration.Section
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using System.Linq;
	using Terrasoft.Common;
	using Terrasoft.Configuration.PageEntity;
	using Terrasoft.Configuration.SspEntity;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Packages;
 	using global::Common.Logging;
	using System.Diagnostics;

	#region Class: SspSectionRepository

	[DefaultBinding(typeof(ISectionRepository), Name = "SSP")]
	public class SspSectionRepository : SectionRepository
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="SysPortal"/> instance unique identifier.
		/// </summary>
		private readonly Guid _sspId = Guid.Parse("C8565240-1DA3-4A68-BD4E-280F17B0D32E");

		/// <summary>
		/// <see cref="ISspEntityRepository"/> implementation instance.
		/// </summary>
		private readonly ISspEntityRepository _sspEntityRepository;

		/// <summary>
		/// <see cref="IResourceStorage"/> implementation instance.
		/// </summary>
		private readonly IResourceStorage _resourceStorage;

		/// <summary>
		/// <see cref="IPageEntityRepository"/> implementation instance.
		/// </summary>
		private readonly IPageEntityRepository _pageEntityRepository;

		/// <summary>
		/// <see cref="DBSecurityEngine"/> implementation instance.
		/// </summary>
		private readonly DBSecurityEngine _dbSecurityEngine;

		/// <summary>
		/// <see cref="PackageElementUtilities"/> implementation instance.
		/// </summary>
		private readonly PackageElementUtilities _packageElementUtilities;

		/// <summary>
 	 	/// <see cref="ILog"/> implementation instance.
 	 	/// </summary>
 	 	private static readonly ILog _log = LogManager.GetLogger("Workplace");
 	 	
 	 	private Guid _sessionId = Guid.NewGuid();

		#endregion

		#region Properties: Protected

		/// <summary>
		/// <see cref="IPageEntityRepository"/> implementation instance.
		/// </summary>
		protected IPageEntityRepository PageEntityRepository {
			get => _pageEntityRepository;
		}

		/// <summary>
		/// <see cref="ISspEntityRepository"/> implementation instance.
		/// </summary>
		protected ISspEntityRepository SspEntityRepository {
			get => _sspEntityRepository;
		}


		#endregion

		#region Constructors: Public

		public SspSectionRepository(UserConnection uc) : base(uc) {
			_sspEntityRepository = ClassFactory.Get<ISspEntityRepository>(new ConstructorArgument("uc", uc));
			_pageEntityRepository = ClassFactory.Get<IPageEntityRepository>(new ConstructorArgument("uc", uc));
			_packageElementUtilities = ClassFactory.Get<PackageElementUtilities>(new ConstructorArgument("userConnection", uc));
			_resourceStorage = uc.ResourceStorage;
			_dbSecurityEngine = uc.DBSecurityEngine;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates section type is not SSP exception message.
		/// </summary>
		/// <param name="sectionId"><see cref="Section"/> unique identifier.</param>
		/// <returns>Section type is not SSP exception message.</returns>
		private string GetSectionNotSspTypeMessage(Guid sectionId) {
			var messageTpl = GetSectionExceptionMessage("SectionNotSspTypeTpl");
			return string.Format(messageTpl, sectionId.ToString());
		}

		/// <summary>
		/// Creates section exception message.
		/// </summary>
		/// <param name="exceptionMessageKey">Localizable string name.</param>
		/// <returns>Section exception message.</returns>
		private string GetSectionExceptionMessage(string exceptionMessageKey) {
			return new LocalizableString(_resourceStorage, "SectionExceptionResources",
				$"LocalizableStrings.{exceptionMessageKey}.Value").ToString();
		}

		/// <summary>
		/// Returns lookup fields reference schemas unique identifiers list used on section page.
		/// </summary>
		/// <param name="sectionMainEntity">Section main entity <see cref="ISchemaManagerItem"/> instance.</param>
		/// <returns>Entities unique identifiers list.</returns>
		private IEnumerable<Guid> GetPageLookupColumnEntityIds(ISchemaManagerItem sectionMainEntity) {
			return GetPageLookupColumnEntityIds(sectionMainEntity.UId);
		}

		/// <summary>
		/// Returns used by <paramref name="section"/> edit pages detail entities.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		/// <returns>Entities unique identifiers list.</returns>
		private IEnumerable<Guid> GetPageDetailsEntityIds(Section section) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetPageDetailsEntityIds", watch, section);
			var pages = _pageEntityRepository.GetAll().Where(p => p.SysModuleEntityId == section.SysModuleEntityId);
			LogDebug($"GetPageDetailsEntityIds PageEntityRepository.GetAll ended", watch);
			List<Guid> result = new List<Guid>();
			if (pages.IsEmpty()) {
			        LogDebug($"GetPageDetailsEntityIds End without select", watch,section);
			        return result;
			}
			var select = new Select(UserConnection)
				.Column("EntitySchemaUId")
				.From("SspPageDetail")
				.Where("CardSchemaUId").In(Column.Parameters(pages.Select(p => p.CardSchemaUId))) as Select;
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						result.AddIfNotExists(
							dataReader.GetColumnValue<Guid>("EntitySchemaUId")
						);
					}
				}
			}
			LogDebug($"GetPageDetailsEntityIds End with select", watch, section);
			return result;
		}


		/// <summary>
		/// Sets <see cref="EntitySchema.IsSSPAvailable"/> property value for <paramref name="section"/>.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void SetIsSSPAvailableProperty(Section section) {
			var sectionMainEntity = GetSectionEntitySchemaItem(section);
			if (sectionMainEntity == null) {
				return;
			}
			_sspEntityRepository.SetEntitySspAllowed(sectionMainEntity.UId);
		}

		/// <summary>
		/// Checks is <paramref name="sysModuleEntityId"/> already registered as <paramref name="sspId"/> section.
		/// </summary>
		/// <param name="sysModuleEntityId">Module entity identifier.</param>
		/// <param name="sspId">Self service portal identifier.</param>
		/// <returns><c>True</c> if <paramref name="section"/> registered as <paramref name="sspId"/> section.
		/// Otherwise returns <c>false</c>.</returns>
		private bool IsRegisteredAsSsp(Guid sysModuleEntityId, Guid sspId) {
			var select = new Select(UserConnection).Top(1)
				.Column("Id")
			.From("SysModuleEntityInPortal")
			.Where("SysModuleEntityId").IsEqual(Column.Parameter(sysModuleEntityId))
			.And("SysPortalId").IsEqual(Column.Parameter(sspId)) as Select;
			return select.ExecuteScalar<Guid>().IsNotEmpty();
		}

		/// <summary>
		/// Fills <see cref="SysModuleEntityInPortal"/> table for <paramref name="sysModuleEntityId"/>.
		/// </summary>
		/// <param name="sysModuleEntityId">Module entity identifier.</param>
		/// <param name="sspId">Self service portal identifier.</param>
		private void FillSysModuleEntityInPortal(Guid sysModuleEntityId, Guid sspId) {
			if (IsRegisteredAsSsp(sysModuleEntityId, sspId)) {
				return;
			}
			var schema = EntitySchemaManager.GetInstanceByName("SysModuleEntityInPortal");
			var entity = schema.CreateEntity(UserConnection);
			entity.SetDefColumnValues();
			entity.SetColumnValue("SysModuleEntityId", sysModuleEntityId);
			entity.SetColumnValue("SysPortalId", sspId);
			entity.Save();
			SavePackageSchemaData("SysModuleEntityInPortal", new List<object>() { entity.PrimaryColumnValue });
		}

		/// <summary>
		/// Fills <see cref="SysModuleEntityInPortal"/> table for <paramref name="section"/>  page details entities.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void FillPageDetailEntitiesSysModuleEntityInPortal(Section section) {
			List<Guid> relatedEntitiesUId = GetPageDetailsEntityIds(section).ToList();
			var sysModuleEntitySchema = EntitySchemaManager.GetInstanceByName("SysModuleEntity");
			var sysModuleEditSchema = EntitySchemaManager.GetInstanceByName("SysModuleEdit");
			var sysModuleEntityInPortalExisting = new List<(bool isNotInPortalExist, Guid sysModuleEntityId, Guid pagesId, Guid sysEntitySchemaUId)>();
			FillSysModuleEditList(sysModuleEntityInPortalExisting, relatedEntitiesUId);
			relatedEntitiesUId.ForEach((relatedEntityUId) => {
				var filteredBySysEntitySchemaUId = sysModuleEntityInPortalExisting.Where((item) => item.sysEntitySchemaUId.Equals(relatedEntityUId));
				var isSysModuleEntityInPortalExist = filteredBySysEntitySchemaUId.Any((item) => !item.isNotInPortalExist);
				if (isSysModuleEntityInPortalExist) {
					return;
				}
				var groupBySysModuleEntity = filteredBySysEntitySchemaUId.Where(item => item.isNotInPortalExist).GroupBy(item => item.sysModuleEntityId);
				groupBySysModuleEntity.ForEach((sysModuleEntityGroup) => {
					var pagesId = sysModuleEntityGroup.Select(item => item.pagesId);
					var entity = sysModuleEntitySchema.CreateEntity(UserConnection);
					var newModuleEntityId = Guid.NewGuid();
					entity.SetDefColumnValues();
					entity.PrimaryColumnValue = newModuleEntityId;
					entity.SetColumnValue("SysEntitySchemaUId", relatedEntityUId);
					entity.Save();
					pagesId.ForEach(sysModuleEditId => {
						var sysModuleEditEntity = sysModuleEditSchema.CreateEntity(UserConnection);
						if (sysModuleEditEntity.FetchFromDB(sysModuleEditId)) {
							var newSysModuleEditEntity = (Entity)sysModuleEditEntity.Clone();
							newSysModuleEditEntity.SetDefColumnValues();
							newSysModuleEditEntity.PrimaryColumnValue = Guid.NewGuid();
							newSysModuleEditEntity.SetColumnValue("SysModuleEntityId", newModuleEntityId);
							newSysModuleEditEntity.Save();
						}
					});
					FillSysModuleEntityInPortal(newModuleEntityId, _sspId);
				});

			});
		}

		/// <summary>
		/// Fills list of existing SysModuleEdit.
		/// </summary>
		/// <param name="list">List for filling.</param>
		/// <param name="relatedEntitiesUId">List of schemas UId.</param>
		private void FillSysModuleEditList(List<(bool, Guid, Guid, Guid)> list, List<Guid> relatedEntitiesUId) {
			if (!relatedEntitiesUId.Any()) {
				return;
			}
			var select = new Select(UserConnection)
				.Column("smen", "Id").As("SysModuleEntityId")
				.Column("smp", "Id").As("SysModuleEntityInPortalId")
				.Column("smen", "SysEntitySchemaUId").As("SysEntitySchemaUId")
				.Column("sme", "Id").As("PagesId")
			.From("SysModuleEdit").As("sme")
				.InnerJoin("SysModuleEntity").As("smen").On("smen", "Id").IsEqual("sme", "SysModuleEntityId")
				.LeftOuterJoin("SysModuleEntityInPortal").As("smp").On("smp", "SysModuleEntityId").IsEqual("smen", "Id")
			.Where("smen", "SysEntitySchemaUId").In(Column.Parameters(relatedEntitiesUId)) as Select;

			select.ExecuteReader(reader => {
				var sysModuleEntityInPortalId = reader.GetColumnValue<Guid>("SysModuleEntityInPortalId");
				var sysModuleEntityId = reader.GetColumnValue<Guid>("SysModuleEntityId");
				var sysEntitySchemaUId = reader.GetColumnValue<Guid>("SysEntitySchemaUId");
				var pagesId = reader.GetColumnValue<Guid>("PagesId");
				list.Add((sysModuleEntityInPortalId.IsEmpty(), sysModuleEntityId, pagesId, sysEntitySchemaUId));
			});
		}

		/// <summary>
		/// Fills <see cref="SysSSPEntitySchemaAccessList"/> table for <paramref name="section"/>.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void FillSysSSPEntitySchemaAccessList(Section section) {
			var usedEntities = GetRelatedEntityIds(section);
			AddSchemaAccessAndSavePackageSchemaData(usedEntities);
		}

		/// <summary>
		/// Returns <see cref="SysPackageSchemaData"/> instance name for <paramref name="entitySchemaName"/>.
		/// </summary>
		/// <param name="entitySchemaName"><see cref="Core.Entities.EntitySchema"/> instance name.</param>
		/// <returns><see cref="SysPackageSchemaData"/> instance name.</returns>
		private string GetPackageSchemaDataName(string entitySchemaName) {
			var id = Guid.NewGuid().ToString().Replace("-", string.Empty);
			return $"{entitySchemaName}_{id}";
		}

		/// <summary>
		/// Saves records from <paramref name="recordIds"/> to <see cref="SysPackageSchemaData"/> instance.
		/// </summary>
		/// <param name="entitySchemaName"><see cref="Core.Entities.EntitySchema"/> instance name.</param>
		/// <param name="recordIds">Record collection identifiers .</param>
		/// <returns><see cref="SysPackageSchemaData"/> instance name.</returns>
		private void SavePackageSchemaData(string entitySchemaName, List<object> recordIds) {
			Guid packageUId = (Guid)Core.Configuration.SysSettings.GetValue(UserConnection, "CurrentPackageId");
			Guid packageId = WorkspaceUtilities.GetPackageIdByUId(packageUId, UserConnection);
			string packageSchemaDataName = GetPackageSchemaDataName(entitySchemaName);
			EntitySchemaQuery query = new EntitySchemaQuery(EntitySchemaManager, entitySchemaName);
			query.AddAllSchemaColumns(true);
			query.Filters.Add(query.CreateFilterWithParameters(FilterComparisonType.Equal,
				query.RootSchema.PrimaryColumn.Name, recordIds.ToArray()));
			_packageElementUtilities.SavePackageSchemaData(packageSchemaDataName, packageId, query, string.Empty);
		}

		/// <summary>
		/// Sets administration by operations for related entities.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void SetEntityEnabledAdministratedByOperations(Section section) {
			var usedEntities = GetRelatedEntityIds(section);
			SetEntityEnabledAdministratedByOperations(usedEntities);
		}

		/// <summary>
		/// Returns <see cref="SysEntitySchemaOperationRight"/> identifiers collection for <paramref name="entityUIds"/>.
		/// collection to result <see cref="SysEntitySchemaOperationRight"/> identifiers collection.
		/// </summary>
		/// <param name="entityUIds"><see cref="Entity"/> instance unique identifier.</param>
		/// <returns><see cref="SysEntitySchemaOperationRight"/> identifiers collection for <paramref name="entityUIds"/>.</returns>
		private List<object> GetAdministratedByOperationsIds(IEnumerable<Guid> entityUIds) {
			var sysAdminUnits = new List<Guid>()
				{UsersConsts.AllEmployersSysAdminUnitUId, UsersConsts.PortalUsersSysAdminUnitUId};
			var select = new Select(UserConnection)
				.Column("Id")
			.From("SysEntitySchemaOperationRight")
			.Where("SubjectSchemaUId").In(Column.Parameters(entityUIds))
				.And("SysAdminUnitId").In(Column.Parameters(sysAdminUnits)) as Select;
			var ids = new List<object>();
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						Guid id = dataReader.GetColumnValue<Guid>("Id");
						ids.Add(id);
					}
				}
			}
			return ids;
		}

		/// <summary>
		/// Returns not administrated by records <see cref="Entity"/> identifier collection from <paramref name="entityIds"/>.
		/// </summary>
		/// <param name="entityIds"><see cref="Entity"/> identifier collection.</param>
		/// <returns><see cref="Entity"/> identifiers collection.</returns>
		private IEnumerable<Guid> GetNonAdministratedByRecordsEntityIds(IEnumerable<Guid> entityIds) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetNonAdministratedByRecordsEntityIds", watch);
			var result = new List<Guid>();
			foreach (var entityId in entityIds) {
				LogDebug($"GetNonAdministratedByRecordsEntityIds start foreach EntitySchemaManager.GetItemByUId {entityId}", watch);
				var schema = EntitySchemaManager.GetItemByUId(entityId);
				LogDebug($"GetNonAdministratedByRecordsEntityIds end foreach EntitySchemaManager.GetItemByUId {entityId}", watch);
				LogDebug($"GetNonAdministratedByRecordsEntityIds start dbSecurityEngine.GetIsEntitySchemaAdministratedByRecords {schema.Name}", watch);
				var schemaIsAdministratedByRecords = _dbSecurityEngine.GetIsEntitySchemaAdministratedByRecords(schema.Name);
				LogDebug($"GetNonAdministratedByRecordsEntityIds end dbSecurityEngine.GetIsEntitySchemaAdministratedByRecords {schema.Name}", watch);
				if (!schemaIsAdministratedByRecords) {
					result.AddIfNotExists(entityId);

				}
			}
			LogDebugEnd("GetNonAdministratedByRecordsEntityIds", watch);
			return result;
		}

		/// <summary>
		/// Returns <see cref="Entity"/> schema captions from <paramref name="entityIds"/>.
		/// </summary>
		/// <param name="entityIds">><see cref="Entity"/> identifier collection.</param>
		/// <returns>><see cref="Entity"/> schema caption collection.</returns>
		private IEnumerable<string> GetEntitySchemaCaptions(IEnumerable<Guid> entityIds) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetEntitySchemaCaptions", watch);
			var captions = new List<string>();
			foreach (var entityId in entityIds) {
				LogDebug($"GetEntitySchemaCaptions start foreach EntitySchemaManager.GetItemByUId {entityId}", watch);
				var item = EntitySchemaManager.FindItemByUId(entityId);
				LogDebug($"GetEntitySchemaCaptions end foreach EntitySchemaManager.GetItemByUId {entityId}", watch);
				if (item != null) {
					captions.AddIfNotExists(item.Caption);
				}
			}
			LogDebugEnd("GetEntitySchemaCaptions", watch);
			return captions;
		}

		/// <summary>
		/// Sets <see cref="Core.Entities.EntitySchema.AdministratedByRecords"/> property
		/// for <paramref name="entityIds"/> <see cref="Core.Entities.EntitySchema"/>.
		/// </summary>
		/// <param name="entityIds">><see cref="Entity"/> identifier collection.</param>
		private void SetSchemasAdministratedByRecords(IEnumerable<Guid> entityIds) {
			var workspaceObjectsSchema = EntitySchemaManager.GetInstanceByName("VwWorkspaceObjects");
			Entity workspaceObject = workspaceObjectsSchema.CreateEntity(UserConnection);
			foreach (var nonAdministratedEntityId in entityIds) {
				var conditions = new Dictionary<string, object> {
					{ "UId", nonAdministratedEntityId },
					{ "SysWorkspace", UserConnection.Workspace.Id}
				};
				if (workspaceObject.FetchFromDB(conditions)) {
					workspaceObject.SetColumnValue("AdministratedByRecords", true);
					workspaceObject.Save();
				}
			}
		}

		/// <summary>
		/// Returns related to <paramref name="section"/> <see cref="Entity"/> schema identifier collection
		/// what are registered as module object and not administrated by records.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		/// <returns><see cref="Entity"/> schema identifier collection.</returns>
		private IEnumerable<Guid> GetSectionNonAdministratedByRecordsEntityIds(Section section) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetSectionNonAdministratedByRecordsEntityIds", watch, section);
			IEnumerable<Guid> relatedEntityIds = GetRelatedEntityIds(section);
			if (relatedEntityIds.IsNullOrEmpty()) {
				LogDebugEnd("GetSectionNonAdministratedByRecordsEntityIds not related", watch, section);
				return new List<Guid>();
			}
			var allSections = GetAll();
			IEnumerable<Guid> moduleEntityIds = relatedEntityIds.Where(
				 (e) => {
					 LogDebug($"GetSectionNonAdministratedByRecordsEntityIds GetAllSections", watch);
					 return allSections.Any(s => s.EntityUId == e && s.SchemaName.IsNotNullOrEmpty());
				 });
			LogDebug($"GetSectionNonAdministratedByRecordsEntityIds Calculation of moduleEntityIds ended", watch);
			if (moduleEntityIds.IsNullOrEmpty()) {
				LogDebugEnd("GetSectionNonAdministratedByRecordsEntityIds not module entity", watch, section);
				return new List<Guid>();
			}
			var result = GetNonAdministratedByRecordsEntityIds(moduleEntityIds);
			LogDebugEnd("GetSectionNonAdministratedByRecordsEntityIds", watch, section);
			return result;
		}
		/// <summary>
		/// Message logging to external storage.
		/// </summary>
		/// <param name="message">Logging message</param>
		/// <param name="stopwatch"><see cref="Stopwatch"/> instance.</param>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void LogDebug(string message, Stopwatch stopwatch, Section section = null) {
			_log.Info($"[{_sessionId}] [{stopwatch.ElapsedMilliseconds}ms]  {message}, {section?.Caption}:{section?.Id}.");
		}


		/// <summary>
		/// Logging starting methods message to external storage.
		/// </summary>
		/// <param name="message">Logging message</param>
		/// <param name="stopwatch"><see cref="Stopwatch"/> instance.</param>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void LogDebugStart(string methodName, Stopwatch stopwatch, Section section = null) {
			LogDebug($"{methodName} Start", stopwatch, section);
		}

		/// <summary>
		/// Logging ended methods message to external storage.
		/// </summary>
		/// <param name="message">Logging message</param>
		/// <param name="stopwatch"><see cref="Stopwatch"/> instance.</param>
		/// <param name="section"><see cref="Section"/> instance.</param>
		private void LogDebugEnd(string methodName, Stopwatch stopwatch, Section section = null) {
			LogDebug($"{methodName} End", stopwatch, section);
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Allows schema access on ssp and saves package data.
		/// </summary>
		/// <param name="usedEntities">Lists of entities identifiers to be enabled on ssp.</param>
		protected void AddSchemaAccessAndSavePackageSchemaData(IEnumerable<Guid> usedEntities) {
			var schema = EntitySchemaManager.GetInstanceByName("SysSSPEntitySchemaAccessList");
			var ids = new List<object>();
			foreach (var entityId in usedEntities) {
				var entity = schema.CreateEntity(UserConnection);
				entity.SetDefColumnValues();
				if (entity.FetchFromDB("EntitySchemaUId", entityId)) {
					continue;
				}
				entity.PrimaryColumnValue = Guid.NewGuid();
				entity.SetColumnValue("EntitySchemaUId", entityId);
				entity.Save();
				ids.Add(entity.PrimaryColumnValue);
			}
			if (ids.IsNotEmpty()) {
				SavePackageSchemaData("SysSSPEntitySchemaAccessList", ids);
			}
		}

		/// <summary>
		/// Returns list of lookup entities from schema.
		/// </summary>
		/// <param name="schemaUId">Entity schema UId.</param>
		/// <returns>List of lookup entities.</returns>
		protected IEnumerable<Guid> GetPageLookupColumnEntityIds(Guid schemaUId) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetPageLookupColumnEntityIds", watch);
			var result = new List<Guid>();
			var sspAvailableColumns = _sspEntityRepository.GetSspColumnAccessList(schemaUId);
			LogDebug($"GetPageLookupColumnEntityIds SspEntityRepository.GetSspColumnAccessList {schemaUId} ended", watch);
			var schema = EntitySchemaManager.GetInstanceByUId(schemaUId);
			LogDebug($"GetPageLookupColumnEntityIds EntitySchemaManager.GetInstanceByUId {schemaUId} ended", watch);
			foreach (var sspEntityColumnUId in sspAvailableColumns) {
				var column = schema.Columns.FindByUId(sspEntityColumnUId);
				LogDebug($"GetPageLookupColumnEntityIds foreach schema.Columns.FindByUId {sspEntityColumnUId} ended", watch);
				if (column != null && column.ReferenceSchemaUId.IsNotEmpty()) {
					result.AddIfNotExists(column.ReferenceSchemaUId);
				}
			}
			LogDebugEnd("GetPageLookupColumnEntityIds", watch);
			return result;
		}

		/// <summary>
		/// Enables administrated by operation for the entities <param name="usedEntities"></param>
		/// </summary>
		/// <param name="usedEntities">List of entities identifiers to be administrated.</param>
		protected void SetEntityEnabledAdministratedByOperations(IEnumerable<Guid> usedEntities) {
			foreach (var entityId in usedEntities) {
				_sspEntityRepository.SetEntityAdministratedByOperations(entityId);
			}
			var ids = new List<object>();
			if (usedEntities.IsNotEmpty()) {
				ids = GetAdministratedByOperationsIds(usedEntities);
			}
			if (ids.IsNotEmpty()) {
				SavePackageSchemaData("SysEntitySchemaOperationRight", ids);
			}
		}

		/// <inheritdoc />
		protected override void SetSectionsByTypeFilters(Select select, SectionType type) {
			if (type == SectionType.SSP) {
				base.SetSectionsByTypeFilters(select, type);
				return;
			}
			select.Where("sm", "Type").IsEqual(Column.Parameter(SectionType.General))
				.And().Not().Exists(
					new Select(UserConnection)
							.Column("sm1", "Id")
						.From("SysModule").As("sm1")
						.InnerJoin("SysModuleEntity").As("sme1").On("sme1", "Id").IsEqual("sm1", "SysModuleEntityId")
						.Where("sme1", "SysEntitySchemaUId").IsEqual("sme", "SysEntitySchemaUId")
						.And("sm1", "Type").IsEqual(Column.Parameter(SectionType.SSP)));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc />
		public override void Save(Section section) {
			if (section.Type != SectionType.SSP) {
				var message = GetSectionNotSspTypeMessage(section.Id);
				throw new SectionNotSspTypeException(message);
			}
			_pageEntityRepository.ClearCache();
			FillSysModuleEntityInPortal(section.SysModuleEntityId, _sspId);
			FillPageDetailEntitiesSysModuleEntityInPortal(section);
			FillSysSSPEntitySchemaAccessList(section);
			SetIsSSPAvailableProperty(section);
			SetEntityEnabledAdministratedByOperations(section);
			ClearCache();
		}

		/// <inheritdoc />
		public override IEnumerable<Guid> GetRelatedEntityIds(Section section) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetRelatedEntityIds", watch);
			var result = base.GetRelatedEntityIds(section) as List<Guid>;
			LogDebug("GetRelatedEntityIds Executing base.GetRelatedEntityIds ended", watch, section);
			var sectionMainEntity = GetSectionEntitySchemaItem(section);
			LogDebug("GetRelatedEntityIds Executing base.GetSectionEntitySchemaItem ended", watch, section);
			if (section.Type != SectionType.SSP || sectionMainEntity == null) {
				LogDebugEnd("GetRelatedEntityIds", watch);
				return result;
			}
			result.AddRangeIfNotExists(GetPageLookupColumnEntityIds(sectionMainEntity));
			result.AddRangeIfNotExists(GetPageDetailsEntityIds(section));
			LogDebugEnd("GetRelatedEntityIds", watch);
			return result;
		}

		/// <inheritdoc />
		public override IEnumerable<string> GetSectionNonAdministratedByRecordsEntityCaptions(Section section) {
			var watch = Stopwatch.StartNew();
			LogDebugStart("GetSectionNonAdministratedByRecordsEntityCaptions", watch, section);
			IEnumerable<Guid> nonAdministratedEntityIds = GetSectionNonAdministratedByRecordsEntityIds(section);
			var result = GetEntitySchemaCaptions(nonAdministratedEntityIds);
			LogDebugEnd("GetSectionNonAdministratedByRecordsEntityCaptions", watch, section);
			return result;
		}

		/// <inheritdoc />
		public override void SetSectionSchemasAdministratedByRecords(Section section) {
			IEnumerable<Guid> nonAdministratedEntityIds = GetSectionNonAdministratedByRecordsEntityIds(section);
			SetSchemasAdministratedByRecords(nonAdministratedEntityIds);
		}

		#endregion

	}

	#endregion

}
