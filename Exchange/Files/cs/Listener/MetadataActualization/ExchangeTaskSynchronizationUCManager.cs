namespace Terrasoft.Configuration
{
	using System;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeTaskSynchronizationUCManager

	/// <summary>
	/// Class provides users that has enabled exchange calendars synchronization list.
	/// </summary>
	[DefaultBinding(typeof(ISynchronizationUCManager), Name = "ExchangeTask")]
	public class ExchangeTaskSynchronizationUCManager : ExchangeCalendarSynchronizationUCManager
	{

		#region Properties: Public

		/// <summary>
		/// <see cref="ISynchronizationUCManager.SynchronizationControllerName"/>
		/// </summary>
		public override string SynchronizationControllerName => "ExchangeTaskMetaDataActualizer";

		#endregion

		/// <summary>
		/// Returns related to <paramref name="entity"/> activity unique identifier.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <returns>Related to <paramref name="entity"/> activity unique identifier.</returns>
		protected override Guid GetRelatedEntityId(Entity entity) {
			switch (entity.SchemaName) {
				case "Activity":
					return entity.PrimaryColumnValue;
				default:
					return Guid.Empty;
			}
		}

		#region Methods: Protected

		/// <summary>
		/// Searches users that can synchronize <paramref name="entityId"/> to exchange tasks list.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="entityId">Activity instance unique identifier.</param>
		/// <returns>Users that can synchronize <paramref name="entityId"/> to exchange tasks list.</returns>
		protected override Select GetUsersSelect(UserConnection userConnection, Guid entityId) {
			if (!CheckActivityFromCalendar(userConnection, entityId, false)) {
				return null;
			}
			return new Select(userConnection)
					.Column("MSS", "CreatedById")
				.From("Activity").As("A")
				.InnerJoin("MailboxSyncSettings").As("MSS").On("A", "OwnerId").IsEqual("MSS", "CreatedById")
				.InnerJoin("ActivitySyncSettings").As("ASS").On("ASS", "MailboxSyncSettingsId").IsEqual("MSS", "Id")
				.Where("A", "Id").IsEqual(Column.Parameter(entityId))
				.And("ASS", "ExportActivities").IsEqual(Column.Parameter(true))
				.And("MSS", "SynchronizationStopped").IsEqual(Column.Parameter(false)) as Select;
		}

		#endregion

	}

	#endregion

}
