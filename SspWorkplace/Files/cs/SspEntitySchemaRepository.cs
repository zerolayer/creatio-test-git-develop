namespace Terrasoft.Configuration.Section
{
	using System;
	using System.Collections.Generic;
	using System.Data;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Factories;

	/// <summary>
	/// Repository for working with access to schemas on ssp.
	/// </summary>
	[DefaultBinding(typeof(ISspEntitySchemaRepository))]
	public class SspEntitySchemaRepository : SspSectionRepository, ISspEntitySchemaRepository
	{

		#region Constructors: Public

		/// <summary>
		/// Creates new instance of <see cref="SspEntitySchemaRepository"/>
		/// </summary>
		/// <param name="userConnection">User connection.</param>
		public SspEntitySchemaRepository(UserConnection userConnection) : base(userConnection) {
		}

		#endregion

		#region Methods: Private

		private void FillSysSSPEntitySchemaAccessList(Guid entitySchemaUId, Guid cardSchemaUId) {
			var usedEntities = GetRelatedEntityIds(entitySchemaUId, cardSchemaUId);
			AddSchemaAccessAndSavePackageSchemaData(usedEntities);
		}

		private IEnumerable<Guid> GetRelatedEntityIds(Guid entitySchemaUId, Guid cardSchemaUId) {
			var result = new List<Guid>();
			result.AddRangeIfNotExists(GetPageLookupColumnEntityIds(entitySchemaUId));
			result.AddRangeIfNotExists(GetPageDetailsEntityIds(cardSchemaUId));
			return result;
		}

		private IEnumerable<Guid> GetPageDetailsEntityIds(Guid cardSchemaUId) {
			List<Guid> result = new List<Guid>();
			var select = new Select(UserConnection)
				.Column("EntitySchemaUId")
				.From("SspPageDetail")
				.Where("CardSchemaUId").IsEqual(Column.Parameter(cardSchemaUId)) as Select;
			using (DBExecutor dbExecutor = UserConnection.EnsureDBConnection()) {
				using (IDataReader dataReader = select.ExecuteReader(dbExecutor)) {
					while (dataReader.Read()) {
						result.AddIfNotExists(
							dataReader.GetColumnValue<Guid>("EntitySchemaUId")
						);
					}
				}
			}
			return result;
		}

		private void SetEntityEnabledAdministratedByOperations(Guid entitySchemaUId, Guid cardSchemaUId) {
			var usedEntities = GetRelatedEntityIds(entitySchemaUId, cardSchemaUId);
			SetEntityEnabledAdministratedByOperations(usedEntities);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Allows access of related schemas by given <param name="entitySchemaUId">schema UId</param> on ssp.
		/// </summary>
		/// <param name="entitySchemaUId">Entity schema UId.</param>
		/// <param name="cardSchemaUId">Card schema UId.</param>
		public void AllowRelatedEntitiesOnSsp(Guid entitySchemaUId, Guid cardSchemaUId) {
			FillSysSSPEntitySchemaAccessList(entitySchemaUId, cardSchemaUId);
			SspEntityRepository.SetEntitySspAllowed(entitySchemaUId);
			SetEntityEnabledAdministratedByOperations(entitySchemaUId, cardSchemaUId);
		}

		#endregion
	}
}
