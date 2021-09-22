namespace Terrasoft.Configuration.Workplace
{
	using System;
	using System.Collections.Generic;

	#region Interface: IWorkplaceRepository

	public interface IWorkplaceRepository
	{

		/// <summary>
		/// Returns all workplaces.
		/// </summary>
		/// <returns><see cref="Workplace"/> collection.</returns>
		IEnumerable<Workplace> GetAll();

		/// <summary>
		/// Saves <paramref name="workplace"/> instance.
		/// </summary>
		/// <param name="workplace"><see cref="Workplace"/> instance.</param>
		/// <returns><c>True</c> if <paramref name="workplace"/> saved. Returns <c>false</c> otherwise.</returns>
		bool SaveWorkplace(Workplace workplace);

		/// <summary>
		/// Removes workplace with <paramref name="workplaceId"/> identifier.
		/// </summary>
		/// <param name="workplaceId"><see cref="Workplace"/> instance unique identifier.</param>
		void DeleteWorkplace(Guid workplaceId);

		/// <summary>
		/// Returns workplace instance with <paramref name="workplaceId"/> identifier.
		/// </summary>
		/// <param name="workplaceId"><see cref="Workplace"/> unique identifier.</param>
		/// <returns><see cref="Workplace"/> instance.</returns>
		Workplace Get(Guid workplaceId);

		/// <summary>
		/// Clears workplace cache records.
		/// </summary>
		void ClearCache();

	}

	#endregion

}
