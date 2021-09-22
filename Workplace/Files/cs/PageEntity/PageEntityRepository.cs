namespace Terrasoft.Configuration.PageEntity
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.Configuration;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Store;

	#region Class PageEntityRepository

	[DefaultBinding(typeof(IPageEntityRepository))]
	public class PageEntityRepository : IPageEntityRepository
	{

		#region Constants: Private

		/// <summary>
		/// Pages <see cref="ICacheStore"/> key.
		/// </summary>
		private const string PagesSessionCacheKey = "All_Pages";

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
		/// <see cref="SysUserInfo"/> instance.
		/// </summary>
		protected readonly SysUserInfo CurrentUser;

		/// <summary>
		/// <see cref="SysWorkspace"/> instance.
		/// </summary>
		protected readonly SysWorkspace Workspace;

		#endregion

		#region Constructors: Public

		public PageEntityRepository(UserConnection uc) {
			UserConnection = uc;
			EntitySchemaManager = uc.EntitySchemaManager;
			ApplicationCache = uc.ApplicationCache;
			CurrentUser = uc.CurrentUser;
			Workspace = uc.Workspace;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Returns <see cref="PageEntity"/> collection from cache.
		/// </summary>
		/// <returns><see cref="PageEntity"/> collection.</returns>
		private List<PageEntity> GetPagesFromCache() {
			return ApplicationCache[PagesSessionCacheKey] as List<PageEntity>;
		}

		/// <summary>
		/// Sets <see cref="PageEntity"/> instance collection to cache.
		/// </summary>
		/// <param name="pages"><see cref="PageEntity"/> instance collection.</param>
		private void SetPagesInCache(List<PageEntity> pages) {
			ApplicationCache[PagesSessionCacheKey] = pages;
		}

		/// <summary>
		/// Creates <see cref="PageEntity"/> data select.
		/// </summary>
		/// <returns><see cref="Select"/> instance.</returns>
		private Select GetClientUnitSchemaNameSelect(string sourceAlias, string sourceColumnAlias) {
			var clientUnitSchemaNameSelect = (Select)new Select(UserConnection)
				.Column("Name")
				.From("VwSysClientUnitSchema")
				.Where("SysWorkspaceId")
				.IsEqual(new QueryParameter("SysWorkspaceId", Workspace.Id, "Guid"))
				.And("UId").IsEqual(sourceAlias, sourceColumnAlias);
			clientUnitSchemaNameSelect.InitializeParameters();
			return clientUnitSchemaNameSelect;
		}

		/// <summary>
		/// Creates new <see cref="PageEntity"/> instance, using information from <paramref name="dataReader"/>.
		/// </summary>
		/// <param name="dataReader"><see cref="IDataReader"/> implementation instance.</param>
		/// <returns><see cref="PageEntity"/> instance.</returns>
		private PageEntity CreatePageInstance(IDataReader dataReader) {
			return PageEntity.CreatePageInstance(dataReader);
		}

		private string GetLczAliasName(string aliasObjectName, string cultureName) {
			var lczAliasName = string.Concat(aliasObjectName, "_", cultureName.Replace("-", ""), "_lcz");
			return lczAliasName.GetHexHashCode();
		}

		private IsNullQueryFunction GetLczColumnQueryFunction(string lczTableName, string lczTableColumnName,
				string mainTableName, string mainTableColumnName) {
			QueryColumnExpression lczTableColumnExpression = Column.SourceColumn(lczTableName, lczTableColumnName);
			QueryColumnExpression mainTableColumnExpression = Column.SourceColumn(mainTableName, mainTableColumnName);
			return Func.IsNull(lczTableColumnExpression, mainTableColumnExpression);
		}

		private void AddLczColumn(Guid cultureId, string cultureName, Select select, string schemaName,
				string schemaAlias, string referencePath, string columnName,
				bool useInnerJoin = true, string columnAlias = null, string tableAlias = null) {
			EntitySchema schema = EntitySchemaManager.GetInstanceByName(schemaName);
			string lczTableName = schema.LocalizationSchemaName;
			string lczTableAliasName = string.IsNullOrEmpty(tableAlias)
				? GetLczAliasName(schemaName, cultureName)
				: tableAlias;
			string lczColumnAliasName = string.IsNullOrEmpty(columnAlias)
				? GetLczAliasName(string.Concat(schemaName, columnName), cultureName)
				: columnAlias;
			string lczColumnName = columnName;
			IsNullQueryFunction lczColumnQueryFunction = GetLczColumnQueryFunction(lczTableAliasName, lczColumnName,
				schemaAlias, columnName);
			select.Column(lczColumnQueryFunction).As(lczColumnAliasName);
			if (useInnerJoin) {
				select.InnerJoin(lczTableName).As(lczTableAliasName)
					.On(lczTableAliasName, "RecordId").IsEqual(schemaAlias, referencePath)
					.And(lczTableAliasName, "SysCultureId").IsEqual(new QueryParameter(cultureId));
			} else {
				select.LeftOuterJoin(lczTableName).As(lczTableAliasName)
					.On(lczTableAliasName, "RecordId").IsEqual(schemaAlias, referencePath)
					.And(lczTableAliasName, "SysCultureId").IsEqual(new QueryParameter(cultureId));
			}
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Selects pages data and creates <see cref="PageEntity"/> collection.
		/// </summary>
		/// <returns><see cref="PageEntity"/> collection.</returns>
		protected List<PageEntity> GetPagesFromDb() {
			var pages = new List<PageEntity>();
			var pagesSelect = GetPagesSelect();
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = pagesSelect.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						pages.AddIfNotExists(CreatePageInstance(dataReader));
					}
				}
			}
			return pages;
		}

		/// <summary>
		/// Creates <see cref="PageEntity"/> data select.
		/// </summary>
		/// <returns>Page <see cref="Select"/> instance.</returns>
		protected Select GetPagesSelect() {
			Guid cultureId = CurrentUser.SysCultureId;
			string cultureName = CurrentUser.SysCultureName;
			var pagesSelect =
				new Select(UserConnection)
					.Column(GetClientUnitSchemaNameSelect("SysModuleEdit", "CardSchemaUId")).As("PageSchemaName")
					.Column(GetClientUnitSchemaNameSelect("SysModule", "CardModuleUId")).As("PageModuleName")
					.Column("SysModuleEdit", "Id").As("SysModuleEditId")
					.Column("SysModuleEdit", "CardSchemaUId").As("CardSchemaUId")
					.Column("SysModuleEdit", "TypeColumnValue").As("TypeColumnValue")
					.Column("SysModuleEdit", "ActionKindName").As("ActionKindName")
					.Column("SysModuleEntity", "Id").As("SysModuleEntityId")
					.Column("SysModuleEntity", "SysEntitySchemaUId").As("SysEntitySchemaUId")
				.From("SysModuleEdit")
				.InnerJoin("SysModuleEntity")
					.On("SysModuleEntity", "Id").IsEqual("SysModuleEdit", "SysModuleEntityId")
				.LeftOuterJoin("SysModule")
					.On("SysModule", "SysModuleEntityId").IsEqual("SysModuleEntity", "Id")
				.Where("SysModuleEdit", "CardSchemaUId").Not().IsNull() as Select;
			AddLczColumn(cultureId, cultureName, pagesSelect,
				"SysModuleEdit", "SysModuleEdit", "Id", "ActionKindCaption", false, "ActionKindCaption",
				"ActionKindCaptionLcz");
			AddLczColumn(cultureId, cultureName, pagesSelect,
				"SysModuleEdit", "SysModuleEdit", "Id", "PageCaption", false, "PageCaption",
				"PageCaptionLcz");
			return pagesSelect;
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IPageEntityRepository.GetAll"/>
		public IEnumerable<PageEntity> GetAll() {
			var pages = GetPagesFromCache();
			if (pages != null) {
				return pages;
			}
			pages = GetPagesFromDb();
			SetPagesInCache(pages);
			return pages;
		}

		/// <inheritdoc cref="IPageEntityRepository.ClearCache"/>
		public void ClearCache() {
			ApplicationCache.Remove(PagesSessionCacheKey);
		}

		#endregion
	}

	#endregion

}