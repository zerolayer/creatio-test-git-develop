namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Sync;

	#region Interface: ISynchronizationUCManager

	public interface ISynchronizationUCManager
	{

		/// <summary>
		/// Related <see cref="ISynchronizationController"/> implementation name.
		/// </summary>
		string SynchronizationControllerName { get; }

		/// <summary>
		/// Returns users that has <paramref name="entity"/> synchronization contact ids list.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <param name="action">Synchronization action for <paramref name="entity"/>.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns>User contacts unique identifiers list.</returns>
		IEnumerable<Guid> GetSynchronizationUsers(Entity entity, SyncAction action, UserConnection userConnection);

		/// <summary>
		/// Returns related entity column value.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns><see cref="EntityColumnValue"/> instance.</returns>
		EntityColumnValue GetParentRelationColumnValue(Entity entity, UserConnection userConnection);

	}

	#endregion

}