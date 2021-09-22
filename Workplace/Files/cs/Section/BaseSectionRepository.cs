namespace Terrasoft.Configuration.Section
{

	using Core.Entities;
	using System;
	using System.Collections.Generic;

	#region Class SectionRepository

	public abstract class BaseSectionRepository: ISectionRepository {

		/// <inheritdoc />
		public abstract Section Get(Guid sectionId);

		/// <inheritdoc />
		public abstract IEnumerable<Section> GetAll();

		/// <inheritdoc />
		public abstract IEnumerable<Section> GetByType(SectionType type);

		/// <inheritdoc />
		public abstract IEnumerable<Guid> GetRelatedEntityIds(Section section);

		/// <inheritdoc />
		public abstract IEnumerable<string> GetSectionNonAdministratedByRecordsEntityCaptions(Section section);

		/// <inheritdoc />
		public abstract void Save(Section section);

		/// <inheritdoc />
		public abstract void ClearCache();

		/// <inheritdoc />
		public virtual void SetSectionSchemasAdministratedByRecords(Section section) { }

	}

	#endregion

}
