namespace Terrasoft.Configuration.Users
{
	using System;
	using System.Collections.Generic;
	using System.Linq;

	#region Class: WorkplaceGroup

	[Serializable]
	public class WorkplaceGroup : IAdministrationUnit
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="SysAdminUnit"/> identifier.
		/// </summary>
		private readonly Guid _groupId;

		/// <summary>
		/// <see cref="SysAdminUnit"/> identifiers collections.
		/// </summary>
		private readonly List<Guid> _userIds;

		#endregion

		#region Constructors: Public

		public WorkplaceGroup(Guid groupId, IEnumerable<Guid> userIds) {
			_groupId = groupId;
			_userIds = userIds.ToList();
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc/>
		public bool GetUnitContainsUser(Guid userId) {
			return _userIds.Contains(userId);
		}

		#endregion
		
	}

	#endregion

}
