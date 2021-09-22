namespace Terrasoft.Configuration.Url
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using global::Common.Logging;
	using Newtonsoft.Json;
	using Terrasoft.Common;
	using Terrasoft.Configuration.Domain;
	using Terrasoft.Configuration.Exception;
	using Terrasoft.Configuration.PageEntity;
	using Terrasoft.Configuration.Section;
	using Terrasoft.Configuration.Workplace;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;

	#region Class: UrlGenerator

	[DefaultBinding(typeof(IUrlGenerator))]
	public class UrlGenerator : IUrlGenerator {

		#region Fiedls: Private

		/// <summary>
		/// Application type identifier.
		/// </summary>
		private readonly Guid _applicationTypeId = Guid.Parse("195785B4-F55A-4E72-ACE3-6480B54C8FA5");

		/// <summary>
		/// <see cref="ILog"/> implementation instance.
		/// </summary>
		private static readonly ILog _log = LogManager.GetLogger("Workplace");

		#endregion

		#region Properties: Protected

		/// <summary>
		/// Page view operation.
		/// </summary>
		protected string Operation {
			get; set;
		} = "edit";

		/// <summary>
		/// <see cref="IDomainResolver"/> implementation instance.
		/// </summary>
		protected IDomainResolver DomainResolver {
			get; private set;
		}

		/// <summary>
		/// <see cref="IPageEntityManager"/> implementation instance.
		/// </summary>
		protected IPageEntityManager PageEntityManager {
			get; private set;
		}

		/// <summary>
		/// <see cref="ISectionManager"/> implementation instance.
		/// </summary>
		protected ISectionManager SectionManager {
			get; private set;
		}

		/// <summary>
		/// <see cref="IWorkplaceManager"/> implementation instance.
		/// </summary>
		protected IWorkplaceManager WorkplaceManager {
			get; private set;
		}

		/// <summary>
		/// <see cref="Core.UserConnection"/> instance.
		/// </summary>
		protected UserConnection UserConnection {
			get; private set;
		}

		/// <summary>
		/// <see cref="EntitySchemaManager"/> instance.
		/// </summary>
		protected EntitySchemaManager EntitySchemaManager {
			get; private set;
		}
		
		#endregion

		#region Constructors: Public

		public UrlGenerator(UserConnection uc) {
			Init(uc);
		}

		public UrlGenerator(UserConnection uc, string operation) {
			Init(uc);
			Operation = operation;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Initialize class dependencies.
		/// </summary>
		/// <param name="uc"></param>
		protected void Init(UserConnection uc) {
			UserConnection = uc;
			DomainResolver = ClassFactory.Get<IDomainResolver>();
			PageEntityManager = ClassFactory.Get<IPageEntityManager>(new ConstructorArgument("uc", UserConnection));
			SectionManager = ClassFactory.Get<ISectionManager>(new ConstructorArgument("uc", UserConnection),
					new ConstructorArgument("sectionType", UserConnection.CurrentUser.ConnectionType.ToString()));
			WorkplaceManager = ClassFactory.Get<IWorkplaceManager>(new ConstructorArgument("uc", UserConnection));
			EntitySchemaManager = UserConnection.EntitySchemaManager;
		}

		/// <summary>
		/// Gets <see cref="PageEntity"/>'s collection, allowed by the current user 
		/// and filtered by entities and type column.
		/// </summary>
		/// <param name="schemaName"><see cref="EntitySchema"/> name.</param>
		/// <param name="recordId">Record unique identifier.</param>
		/// <returns>Return <see cref="PageEntity"/>'s collection</returns>
		private IEnumerable<PageEntity> GetPageEntities(string schemaName, Guid recordId) {
			_log.Info("GetPageEntities Started");
			var entitySchema = EntitySchemaManager.GetInstanceByName(schemaName);
			var sections = GetSections(entitySchema);
			LogInfoObjectList("Allowed entity sections in allowed workplaces", sections);
			var pageEntities = new List<PageEntity>();
			foreach (var section in sections) {
				var sectionPages = PageEntityManager.GetSectionPages(section.Id);
				LogInfoObjectList($"Section '{section.SchemaName}' pages", sectionPages);
				if (section.IsTyped(PageEntityManager)) {
					var typeColumnValue = GetEntityTypeColumnValue(entitySchema, recordId, section.TypeColumnName);
					LogInfoObject($"EntitySchemaName = {entitySchema.Name}, TypecolumnValue is", typeColumnValue);
					sectionPages = FilterPagesByType(sectionPages, typeColumnValue);
				}
				LogInfoObjectList($"Section '{section.SchemaName}' filtered pages", sectionPages);
				pageEntities.AddRangeIfNotExists(sectionPages);
			}
			if (pageEntities.IsEmpty()) {
				_log.Info($"Get pages by schema '{entitySchema.Name}' Uid '{entitySchema.UId}'");
				var entityPages = PageEntityManager.GetEntityPages(entitySchema.UId);
				LogInfoObjectList($"Entity pages", entityPages);
				entityPages = entityPages.Where(p => !p.HasSection);
				LogInfoObjectList($"Entity pages with no sections", entityPages);
				pageEntities.AddRangeIfNotExists(entityPages);
			}
			if (pageEntities.IsEmpty()) {
				throw new NotFoundPageEntityException("PageEntity not found.");
			}
			return pageEntities.OrderByDescending(p => p.TypeColumnValue);
		}

		/// <summary>
		/// Gets <see cref="Section"/>'s allowed by the current user and filtered by entities.
		/// </summary>
		/// <param name="entitySchema"><see cref="EntitySchema"/> instance.</param>
		/// <returns>Return <see cref="Section"/>'s allowed by the current user and filtered by entities.</returns>
		private IEnumerable<Section> GetSections(EntitySchema entitySchema) {
			var currentUserWorkplaces = GetCurrentUserWorkplaces();
			LogInfoObjectList("Allowed workplaces", currentUserWorkplaces);
			var workplacesSectionsIds = GetWorkplacesSectionIds(currentUserWorkplaces);
			var sections = SectionManager.GetSectionsByEntityUId(entitySchema.UId);
			LogInfoObjectList($"Allowed schema '{entitySchema.Name}' sections", sections);
			var sectionType = GetSectionType();
			return sections.Where(s => workplacesSectionsIds.Contains(s.Id) && s.Type == sectionType);
		}

		/// <summary>
		/// Logging <see cref="IEnumerable"/> list of objects.
		/// </summary>
		/// <param name="message">Logging message.</param>
		/// <param name="list">Logging objects list.</param>
		private void LogInfoObjectList(string message, IEnumerable<object> list) {
			foreach (var obj in list) {
				LogInfoObject(message, obj);
			}
		}

		/// <summary>
		/// Logging serialize <paramref name="obj"/> with <paramref name="message"/>.
		/// </summary>
		/// <param name="message">Logging message.</param>
		/// <param name="obj">Logging object.</param>
		private void LogInfoObject(string message, object obj) {
			_log.Info($"{message}. Object '{obj.GetType().Name}': {JsonConvert.SerializeObject(obj)}");
		}

		/// <summary>
		/// Gets section type by current user.
		/// </summary>
		/// <returns>Return section type.</returns>
		private SectionType GetSectionType() {
			return UserConnection.CurrentUser.ConnectionType == UserType.General
				? SectionType.General
				: SectionType.SSP;
		}

		/// <summary>
		/// Gets entity type column value.
		/// </summary>
		/// <param name="entitySchema"<see cref="EntitySchema"/> instance.</param>
		/// <param name="recordId">Record unique identifier.</param>
		/// <param name="typeColumnName">Entity type column name.</param>
		/// <returns>Return entity type column value.</returns>
		private Guid GetEntityTypeColumnValue(EntitySchema entitySchema, Guid recordId, string typeColumnName) {
			if (string.IsNullOrEmpty(typeColumnName)) {
				return Guid.Empty;
			}
			var typeColumn = entitySchema.GetSchemaColumnByPath(typeColumnName);
			var select = new Select(UserConnection)
							.Column(typeColumn.ColumnValueName)
						.From(entitySchema.Name)
						.Where(entitySchema.PrimaryColumn.Name).IsEqual(Column.Parameter(recordId)) as Select;
			try {
				return select.ExecuteScalar<Guid>();
			} catch (Exception ex) {
				throw new NotFoundEntityException($"Entity not found with name = {entitySchema.Name}, recordId = {recordId}", ex);
			}
		}

		/// <summary>
		/// Gets filtered <see cref="PageEntity"/> collection by type, 
		/// if <paramref name="typeColumnValue"/> is not empty.
		/// </summary>
		/// <param name="pages"><see cref="PageEntity"/> collection.</param>
		/// <param name="typeColumnValue">Type column value.</param>
		/// <returns>Return filtered <see cref="PageEntity"/> collection by type.</returns>
		private IEnumerable<PageEntity> FilterPagesByType(IEnumerable<PageEntity> pages, Guid typeColumnValue) {
			if (!typeColumnValue.IsEmpty()) {
				pages = pages.Where(p => p.TypeColumnValue.Equals(typeColumnValue));
			}
			return pages;
		}

		/// <summary>
		/// Get all <see cref="Workplace"/>'s alowed by current user.
		/// </summary>
		/// <returns>Return <see cref="Workplace"/>'s alowed by current user.</returns>
		private IEnumerable<Workplace> GetCurrentUserWorkplaces() {
			var currentUserWorkplaces = WorkplaceManager.GetCurrentUserWorkplaces(_applicationTypeId);
			if (currentUserWorkplaces.IsEmpty()) {
				throw new NotFoundPageEntityException("PageEntity not found, because workpaleces not found.");
			}
			return currentUserWorkplaces;
		}

		/// <summary>
		/// Gets all <see cref="Section"/>'s identifiers for <paramref name="allowedWorkplaces"/>.
		/// </summary>
		/// <param name="allowedWorkplaces"><see cref="Workplace"/> collection.</param>
		/// <returns>Return <see cref="Section"/>'s identifiers collection.</returns>
		private IEnumerable<Guid> GetWorkplacesSectionIds(IEnumerable<Workplace> allowedWorkplaces) {
			var sectionsIds = new List<Guid>();
			allowedWorkplaces.ForEach(w => sectionsIds.AddRangeIfNotExists(w.GetSectionIds()));
			return sectionsIds;
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Gets last segment of Url.
		/// </summary>
		/// <param name="page"><see cref="PageEntity"/> instance.</param>
		/// <param name="recordId">Record unique identifier.</param>
		/// <param name="operation">Page view operation.</param>
		/// <returns>Return last Url segment.</returns>
		protected string GetUrlHash(PageEntity page, Guid recordId, string operation) {
			var urlComponents = new string[] { page.PageModuleName, page.PageSchemaName, operation, recordId.ToString() };
			return string.Join("/", urlComponents);
		}

		/// <summary>
		/// Gets <see cref="PageEntity"/> defining hash of Url.
		/// </summary>
		/// <param name="schemaName"><see cref="EntitySchema"/> name.</param>
		/// <param name="recordId">Record unique identifier.</param>
		/// <returns><see cref="PageEntity"/> instance.</returns>
		protected PageEntity GetPageEntity(string schemaName, Guid recordId) {
			var pageEntities = GetPageEntities(schemaName, recordId);
			LogInfoObjectList($"Final pages", pageEntities);
			return pageEntities.First();
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IUrlGenerator.GetUrl"/>
		public string GetUrl(string schemaName, Guid recordId) {
			try {
				_log.Info("Generate URL Started");
				var defaultUrl = GetDefaultUrl();
				var pageEntity = GetPageEntity(schemaName, recordId);
				LogInfoObject($"Final page", pageEntity);
				var urlHash = GetUrlHash(pageEntity, recordId, Operation);
				var url = string.Concat(defaultUrl, "#", urlHash);
				_log.Info($"Generate URL ended. Url is '{url}'");
				return url;
			} catch (Exception e) {
				_log.Info($"Generate default URL. Error - {e.Message}");
				return GetDefaultUrl();
			}
		}

		/// <inheritdoc cref="IUrlGenerator.GeDefaulttUrl"/>
		public string GetDefaultUrl() {
			var domain = DomainResolver.GetDomain();
			return $"{domain}/Nui/ViewModule.aspx";
		}

		#endregion

	}

	#endregion

}
