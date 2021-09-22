namespace IntegrationV2.MailboxDomain.Repository
{
	using Terrasoft.Core;
	using Terrasoft.Core.Store;

	#region Class: BaseRepository

	/// <summary>
	/// Base repository implementation.
	/// </summary>
	internal abstract class BaseRepository
	{

		#region Fields: Protected

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		protected UserConnection UserConnection;

		/// <summary>
		/// Repository cache name.
		/// </summary>
		protected string CacheName;

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns repository cache.
		/// </summary>
		/// <returns>Repository cache.</returns>
		protected object GetCache() {
			ICacheStore applicationCache = UserConnection.ApplicationCache;
			object store = applicationCache[CacheName];
			return store;
		}

		/// <summary>
		/// Sets value in repository cache.
		/// </summary>
		protected void SetCache(object value) {
			ICacheStore applicationCache = UserConnection.ApplicationCache;
			applicationCache[CacheName] = value;
		}

		#endregion

	}

	#endregion

}
