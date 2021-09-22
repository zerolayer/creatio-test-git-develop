namespace Terrasoft.Configuration.Section
{
    using global::Common.Logging;
    using System;
	using System.Collections.Generic;
	using System.Data;
    using System.Diagnostics;
    using System.Linq;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Store;

	#region Class SectionRepository

	[DefaultBinding(typeof(ISectionRepository), Name = "General")]
	public class SectionRepository : BaseSectionRepository
	{

		#region Fileds: Private

		private readonly string[] _sectionRelatedEntitySufixes = { "File", "Folder", "InFolder", "Tag", "InTag" };

		/// <summary>
		/// <see cref="ILog"/> implementation instance.
		/// </summary>
		private static readonly ILog _log = LogManager.GetLogger("Workplace");

		#endregion

		#region Fields: Protected

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		protected readonly UserConnection UserConnection;

		/// <summary>
		/// <see cref="EntitySchemaManager"/> instance.
		/// </summary>
		protected readonly EntitySchemaManager EntitySchemaManager;

		/// <summary>
		/// <see cref="ICacheStore"/> implementation instance.
		/// Represents application level cache.
		/// </summary>
		protected readonly ICacheStore ApplicationCache;

		/// <summary>
		/// <see cref="IResourceStorage"/> implementation instance.
		/// </summary>
		protected readonly IResourceStorage ResourceStorage;

		#endregion

		#region Constructors: Public

		public SectionRepository(UserConnection uc) {
			UserConnection = uc;
			EntitySchemaManager = uc.EntitySchemaManager;
			ApplicationCache = uc.ApplicationCache;
			ResourceStorage = uc.ResourceStorage;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates <see cref="SysSchema"/> name subselect, using related columns.
		/// </summary>
		/// <param name="sourceAlias">Related table name.</param>
		/// <param name="sourceColumnAlias">SysSchema column name in related table.</param>
		/// <returns><see cref="SysSchema"/> name subselect.</returns>
		private Select GetClientUnitSchemaNameSelect(string sourceAlias, string sourceColumnAlias) {
			var clientUnitSchemaNameSelect = (Select)new Select(UserConnection)
					.Column("Name")
				.From("VwSysClientUnitSchema")
				.Where("SysWorkspaceId")
					.IsEqual(new QueryParameter("SysWorkspaceId", UserConnection.Workspace.Id, "Guid"))
				.And("UId").IsEqual(sourceAlias, sourceColumnAlias);
			clientUnitSchemaNameSelect.InitializeParameters();
			return clientUnitSchemaNameSelect;
		}

		/// <summary>
		/// Creates new <see cref="Section"/> instance, using information from <paramref name="dataReader"/>.
		/// </summary>
		/// <param name="dataReader"><see cref="IDataReader"/> implementation instance.</param>
		/// <returns><see cref="Section"/> instance.</returns>
		private Section CreateSectionInstance(IDataReader dataReader) {
			Guid sectionId = dataReader.GetColumnValue<Guid>("Id");
			int type = dataReader.GetColumnValue<int>("Type");
			string caption = dataReader.GetColumnValue<string>("Caption");
			string schemaName = dataReader.GetColumnValue<string>("SectionSchema");
			string moduleName = dataReader.GetColumnValue<string>("SectionModuleSchema");
			string sectionSchemaName = (string.IsNullOrEmpty(schemaName) && !string.IsNullOrEmpty(moduleName)) ? moduleName : schemaName;
			string typeColumnName = GetTypeColumnName(dataReader, sectionSchemaName);
			Guid sysModuleEntityId = dataReader.GetColumnValue<Guid>("SysModuleEntityId");
			Guid entityUId = dataReader.GetColumnValue<Guid>("EntityUId");
			Guid sysModuleVisaEntityUId = dataReader.GetColumnValue<Guid>("VisaSchemaUId");
			var sectionInWorkplaces = GetSectionWorkplaceIds(sectionId);
			var section = new Section(sectionId, sysModuleEntityId, (SectionType)type) {
				Caption = caption,
				Code = dataReader.GetColumnValue<string>("Code"),
				SchemaName = sectionSchemaName,
				EntityUId = entityUId,
				TypeColumnName = typeColumnName,
				SysModuleVisaEntityUId = sysModuleVisaEntityUId
			};
			section.AddSectionInWorkplaceRange(sectionInWorkplaces);
			return section;
		}

		/// <summary>
		/// Gets type column name.
		/// </summary>
		/// <param name="dataReader"><see cref="IDataReader"/> implementation instance.</param>
		/// <param name="schemaName">Section schema name.</param>
		/// <returns></returns>
		private string GetTypeColumnName(IDataReader dataReader, string schemaName) {
			return dataReader.GetColumnValue<string>("Attribute");
		}

		/// <summary>
		/// Selects <paramref name="sectionId"/> workplaces info.
		/// </summary>
		/// <param name="sectionId"><see cref="Section"/> unique identifier.</param>
		/// <returns>Section workplaces identifiers collection.</returns>
		private IEnumerable<Guid> GetSectionWorkplaceIds(Guid sectionId) {
			var select = new Select(UserConnection)
					.Column("smiw", "Position")
					.Column("smiw", "SysWorkplaceId")
				.From("SysModuleInWorkplace").As("smiw")
				.Where("smiw", "SysModuleId").IsEqual(Column.Parameter(sectionId)) as Select;
			List<KeyValuePair<int, Guid>> result = new List<KeyValuePair<int, Guid>>();
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						result.AddIfNotExists( new KeyValuePair<int, Guid> (
							dataReader.GetColumnValue<int>("Position"),
							dataReader.GetColumnValue<Guid>("SysWorkplaceId")));
					}
				}
			}
			return result.OrderBy(kvp => kvp.Key).Select(kvp => kvp.Value);
		}

		/// <summary>
		/// Returns <see cref="Section"/> collection from cache.
		/// </summary>
		/// <param name="key">Cache key.</param>
		/// <returns><see cref="Section"/> collection.</returns>
		private List<Section> GetFromCache(string key) {
			var sections = ApplicationCache[key] as List<Section>;
			if (sections == null || !sections.Any()) {
				return null;
			}
			return sections;
		}

		/// <summary>
		/// Sets <paramref name="value"/> to cache in <paramref name="key"/>.
		/// </summary>
		/// <param name="key">Cache key.</param>
		/// <param name="value">Cache item value.</param>
		private void SetInCache(string key, List<Section> value) {
			ApplicationCache[key] = value;
		}

		/// <summary>
		/// Clears all related to <paramref name="workplaceId"/> cache items.
		/// </summary>
		private void ClearSectionCache() {
			ApplicationCache.Remove(GetAllCacheKey());
			ApplicationCache.Remove(GetSectionsByTypeKey(SectionType.General));
			ApplicationCache.Remove(GetSectionsByTypeKey(SectionType.SSP));
		}

		/// <summary>
		/// Returns all sections cache key.
		/// </summary>
		/// <returns>All sections cache key.</returns>
		private string GetAllCacheKey() {
			return "All_Sections";
		}

		/// <summary>
		/// When <paramref name="item"/> is not null, adds <paramref name="item"/> unique identifier to <paramref name="list"/>.
		/// </summary>
		/// <param name="list"><see cref="List{Guid}"/> instance.</param>
		/// <param name="item"><see cref="ISchemaManagerItem"/> instance.</param>
		private void AddUIdIfNotNull(List<Guid> list, ISchemaManagerItem item) {
			if (item != null) {
				list.AddIfNotExists(item.UId);
			}
		}

		/// <summary>
		/// Returns section required entities unique identifiers list.
		/// </summary>
		/// <param name="sectionMainEntity">Section main entity <see cref="ISchemaManagerItem"/> instance.</param>
		/// <returns>Section required entities unique identifiers list.</returns>
		private IEnumerable<Guid> GetSectionRequiredEntityIds(ISchemaManagerItem sectionMainEntity) {
			var result = new List<Guid>();
			result.AddIfNotExists(sectionMainEntity.UId);
			foreach (var entityNameSuffix in _sectionRelatedEntitySufixes) {
				AddUIdIfNotNull(result, EntitySchemaManager.FindItemByName(sectionMainEntity.Name + entityNameSuffix));
			}
			return result;
		}

		/// <summary>
		/// Creates section not found exception message.
		/// </summary>
		/// <param name="sectionId"><see cref="Section"/> unique identifier.</param>
		/// <returns>Section not found exception message.</returns>
		private string GetItemNotFoundMessage(Guid sectionId) {
			var messageTpl = new LocalizableString(ResourceStorage, "SectionExceptionResources",
					"LocalizableStrings.SectionNotFoundByIdTpl.Value").ToString();
			return string.Format(messageTpl, sectionId.ToString());
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Creates <see cref="Section"/> data select.
		/// </summary>
		/// <returns><see cref="Select"/> instance.</returns>
		protected Select GetSectionsSelect() {
			Select sectionSchemaSelect = GetClientUnitSchemaNameSelect("sm", "SectionSchemaUId");
			Select sectionModuleSchemaSelect = GetClientUnitSchemaNameSelect("sm", "SectionModuleSchemaUId");
			var select = new Select(UserConnection)
					.Column("sm", "Id")
					.Column("sm", "Caption")
					.Column("sm", "Type")
					.Column("sm", "Code")
					.Column("sm", "Attribute")
					.Column("sm", "SysModuleEntityId")
					.Column("smv", "VisaSchemaUId")
					.Column(sectionSchemaSelect).As("SectionSchema")
					.Column(sectionModuleSchemaSelect).As("SectionModuleSchema")
					.Column("sme", "SysEntitySchemaUId").As("EntityUId")
					.Column("sme", "TypeColumnUId").As("TypeColumnUId")
				.From("SysModule").As("sm")
				.InnerJoin("SysModuleEntity").As("sme")
					.On("sme", "Id").IsEqual("sm", "SysModuleEntityId")
				.LeftOuterJoin("SysModuleVisa").As("smv")
					.On("sm", "SysModuleVisaId").IsEqual("smv", "Id") as Select;
			return select;
		}

		/// <summary>
		/// Selects sections data using <paramref name="select"/> and creates <see cref="Section"/> collection.
		/// If cached result avaliable, select will be skipped.
		/// </summary>
		/// <param name="select"><see cref="Select"/> instance.</param>
		/// <param name="cacheKey">Session cache item key.</param>
		/// <returns><see cref="Section"/> collection.</returns>
		protected List<Section> GetSections(Select select, string cacheKey) {
			var watch = Stopwatch.StartNew();
			_log.Debug($"[GetSections] [{watch.ElapsedMilliseconds}ms] Start.");
			List<Section> sections = GetFromCache(cacheKey);
			if (sections != null) {
				_log.Debug($"[GetSections] [{watch.ElapsedMilliseconds}ms] End from cache.");
				return sections;
			}
			sections = GetSectionsFromDb(select);
			_log.Debug($"[GetSections] [{watch.ElapsedMilliseconds}ms] End from DB.");
			SetInCache(cacheKey, sections);
			_log.Debug($"[GetSections] [{watch.ElapsedMilliseconds}ms] Set to cache.");
			return sections;
		}

		/// <summary>
		/// Selects sections data using <paramref name="select"/> and creates <see cref="Section"/> collection.
		/// </summary>
		/// <param name="select"><see cref="Select"/> instance.</param>
		/// <returns><see cref="Section"/> collection.</returns>
		protected List<Section> GetSectionsFromDb(Select select) {
			List<Section> sections = new List<Section>();
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						sections.AddIfNotExists(CreateSectionInstance(dataReader));
					}
				}
			}
			return sections;
		}

		/// <summary>
		/// Returns <see cref="ISchemaManagerItem"/> used by <paramref name="section"/>.
		/// </summary>
		/// <param name="section"><see cref="Section"/> instance.</param>
		/// <returns><see cref="ISchemaManagerItem"/> used by <paramref name="section"/>.</returns>
		protected ISchemaManagerItem GetSectionEntitySchemaItem(Section section) {
			return EntitySchemaManager.FindItemByUId(section.EntityUId);
		}

		/// <summary>
		/// Returns sections with type general cache key.
		/// </summary>
		/// <returns>Sections with type general cache key.</returns>
		protected string GetSectionsByTypeKey(SectionType type) {
			return $"Sections_{type}";
		}

		/// <summary>
		/// Sets sections by type filters to <paramref name="select"/>.
		/// </summary>
		/// <param name="select"><see cref="Select"/> instance.</param>
		/// <param name="type">Type filter value/</param>
		protected virtual void SetSectionsByTypeFilters(Select select, SectionType type) {
			select.Where("sm", "SysEntitySchemaUId").IsEqual(Column.Parameter(type));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc />
		public override Section Get(Guid sectionId) {
			var select = GetSectionsSelect();
			var cacheKey = GetAllCacheKey();
			var sections = GetSections(select, cacheKey);
			if (sections.Any()) {
				return sections.First(s => s.Id.Equals(sectionId));
			}
			var message = GetItemNotFoundMessage(sectionId);
			throw new ItemNotFoundException(message);
		}

		/// <inheritdoc />
		public override IEnumerable<Section> GetAll() {
			var select = GetSectionsSelect();
			var cacheKey = GetAllCacheKey();
			return GetSections(select, cacheKey);
		}

		/// <inheritdoc />
		public override IEnumerable<Section> GetByType(SectionType type) {
			var select = GetSectionsSelect();
			SetSectionsByTypeFilters(select, type);
			var cacheKey = GetSectionsByTypeKey(type);
			return GetSections(select, cacheKey);
		}

		/// <inheritdoc />
		public override IEnumerable<Guid> GetRelatedEntityIds(Section section) {
			var result = new List<Guid>();
			var sectionMainEntity = GetSectionEntitySchemaItem(section);
			if (sectionMainEntity == null) {
				return result;
			}
			result.AddRangeIfNotExists(GetSectionRequiredEntityIds(sectionMainEntity));
			if (section.SysModuleVisaEntityUId.IsNotEmpty()) {
				AddUIdIfNotNull(result, EntitySchemaManager.FindItemByUId(section.SysModuleVisaEntityUId));
			}
			return result;
		}

		/// <inheritdoc />
		public override IEnumerable<string> GetSectionNonAdministratedByRecordsEntityCaptions(Section section) {
			return new List<string>();
		}

		/// <inheritdoc />
		public override void Save(Section section) {
			ClearCache();
		}

		/// <inheritdoc />
		public override void ClearCache() {
			ClearSectionCache();
		}

		#endregion
	}

	#endregion

}