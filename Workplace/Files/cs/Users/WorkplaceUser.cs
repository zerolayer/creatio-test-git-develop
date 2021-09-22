namespace Terrasoft.Configuration.Users
{
	using System;

	#region Class: WorkplaceUser

	[Serializable]
	public class WorkplaceUser : IAdministrationUnit
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="SysAdminUnit"/> identifier.
		/// </summary>
		private readonly Guid _userId;

		#endregion

		#region Constructors: Public

		public WorkplaceUser(Guid userId) {
			_userId = userId;
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc/>
		public bool GetUnitContainsUser(Guid userId) {
			return _userId.Equals(userId);
		}

		#endregion

	}

	#endregion

}
