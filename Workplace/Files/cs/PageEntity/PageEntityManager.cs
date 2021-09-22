namespace Terrasoft.Configuration.PageEntity
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Configuration.Section;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;

	#region Class: PageEntityManager

	[DefaultBinding(typeof(IPageEntityManager))]
	public class PageEntityManager : IPageEntityManager
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="IPageEntityRepository"/> implementation instance.
		/// </summary>
		private readonly IPageEntityRepository _pageEntityRepository;

		/// <summary>
		/// <see cref="ISectionRepository"/> implementation instance.
		/// </summary>
		private readonly ISectionRepository _sectionRepository;

		#endregion

		#region Constructors: Public

		public PageEntityManager(UserConnection uc) {
			_pageEntityRepository = ClassFactory.Get<IPageEntityRepository>(new ConstructorArgument("uc", uc));
			_sectionRepository = ClassFactory.Get<ISectionRepository>("General", new ConstructorArgument("uc", uc));
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IPageEntityManager.GetSectionPages"/>
		public IEnumerable<PageEntity> GetSectionPages(Guid sectionId) {
			var pages = _pageEntityRepository.GetAll();
			Section pageSection = _sectionRepository.Get(sectionId);
			return pages.Where(p => p.SysModuleEntityId.Equals(pageSection.SysModuleEntityId));
		}

		/// <inheritdoc cref="IPageEntityManager.GetEntityPages"/>
		public IEnumerable<PageEntity> GetEntityPages(Guid entityUId) {
			var pages = _pageEntityRepository.GetAll();
			return pages.Where(p => p.SysEntitySchemaUId.Equals(entityUId));
		}

		#endregion

	}

	#endregion

}