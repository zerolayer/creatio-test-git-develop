namespace Terrasoft.Configuration.SspEntity
{
	using System;
	using System.Collections.Generic;
	using Terrasoft.Core.Entities;

	#region Interface: ISspEntityRepository

	public interface ISspEntityRepository
	{

		#region Methods: Public

		/// <summary>
		/// todo
		/// </summary>
		/// <param name="entityUId"><see cref="Entity"/> unique identifier.</param>
		/// <returns><see cref="SspEntity"/> instance.</returns>
		IEnumerable<Guid> GetSspColumnAccessList(Guid entityUId);

		/// <summary>
		/// Checks that <paramref name="entityName"/> allowed for self service portal.
		/// </summary>
		/// <param name="entityName"><see cref="Entity"/> name.</param>
		/// <returns><c>True</c> if <paramref name="entityName"/> allowed for ssp. Otherwise returns <c>false</c>.
		/// </returns>
		bool IsEntitySspAllowed(string entityName);

		/// <summary>
		/// Sets <see cref="EntitySchema.IsSSPAvailable"/> property for <paramref name="entityUId"/>.
		/// </summary>
		/// <param name="entityUId"><see cref="EntitySchema"/> unique identifier.</param>
		void SetEntitySspAllowed(Guid entityUId);

		/// <summary>
		/// Sets <see cref="EntitySchema.AdministratedByOperations"/> property for <paramref name="entityUId"/>.
		/// </summary>
		/// <param name="entityUId"><see cref="EntitySchema"/> unique identifier.</param>
		void SetEntityAdministratedByOperations(Guid entityUId);

		#endregion

	}

	#endregion

}