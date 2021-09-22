namespace Terrasoft.Configuration
{
	using System;
	using Terrasoft.Common;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeCalendarSynchronizationUCManager

	/// <summary>
	/// Class provides users that has enabled exchange calendars synchronization list.
	/// </summary>
	[DefaultBinding(typeof(ISynchronizationUCManager), Name = "ExchangeCalendar")]
	public class ExchangeCalendarSynchronizationUCManager : ExchangeSynchronizationUCManager
	{

		#region Properties: Public

		/// <summary>
		/// <see cref="ISynchronizationUCManager.SynchronizationControllerName"/>
		/// </summary>
		public override string SynchronizationControllerName => "ExchangeCalendarMetaDataActualizer";

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns related to <paramref name="entity"/> parent column name.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <returns>Related to <paramref name="entity"/> activity unique identifier.</returns>
		protected override string GetRelatedEntityColumnName(Entity entity) {
			switch (entity.SchemaName) {
				case "Activity":
					return entity.Schema.GetPrimaryColumnName();
				case "ActivityParticipant":
					return "ActivityId";
				default:
					return string.Empty;
			}
		}

		/// <summary>
		/// Checks tham <paramref name="entityId"/> activity record showed in calendar.
		/// <paramref name="isFromCalendar"/> parameter contains expected value.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="entityId">Activity recore unique identifier.</param>
		/// <param name="isFromCalendar">Expected Activity ShowInScheduler column value.</param>
		/// <returns><c>True</c> if Activity ShowInScheduler column value equaals <paramref name="isFromCalendar"/>.
		/// Returns <c>false</c> otherwise.</returns>
		protected bool CheckActivityFromCalendar(UserConnection userConnection, Guid entityId, bool isFromCalendar) {
			var select = new Select(userConnection).Top(1)
				.Column("Id")
				.From("Activity")
				.Where("Id").IsEqual(Column.Parameter(entityId))
				.And("TypeId").IsNotEqual(Column.Parameter(ActivityConsts.EmailTypeUId))
				.And("ShowInScheduler").IsEqual(Column.Parameter(isFromCalendar)) as Select;
			return select.ExecuteScalar<Guid>().IsNotEmpty();
		}

		/// <summary>
		/// Searches users that can synchronize <paramref name="entityId"/> to exchange calendar list.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="entityId">Activity instance unique identifier.</param>
		/// <returns>Users that can synchronize <paramref name="entityId"/> to exchange calendar list.</returns>
		protected override Select GetUsersSelect(UserConnection userConnection, Guid entityId) {
			if (!CheckActivityFromCalendar(userConnection, entityId, true)) {
				return null;
			}
			return new Select(userConnection)
					.Column("MSS", "CreatedById")
				.From("ActivityParticipant").As("AP")
				.InnerJoin("MailboxSyncSettings").As("MSS").On("AP", "ParticipantId").IsEqual("MSS", "CreatedById")
				.InnerJoin("ActivitySyncSettings").As("ASS").On("ASS", "MailboxSyncSettingsId").IsEqual("MSS", "Id")
				.Where("AP", "ActivityId").IsEqual(Column.Parameter(entityId))
				.And("ASS", "ExportActivities").IsEqual(Column.Parameter(true))
				.And("MSS", "SynchronizationStopped").IsEqual(Column.Parameter(false)) as Select;
		}

		#endregion

	}

	#endregion

}
