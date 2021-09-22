namespace Terrasoft.Configuration.PageEntity
{

	using System.Collections.Generic;
	
	#region Interface: IPageEntityRepository

	public interface IPageEntityRepository
	{

		/// <summary>
		/// Returns all <see cref="PageEntity"/> collection.
		/// </summary>
		/// <returns>All <see cref="PageEntity"/> instance collection.</returns>
		IEnumerable<PageEntity> GetAll();

		/// <summary>
		/// Clears page repository cache.
		/// </summary>
		void ClearCache();

	}

	#endregion

}