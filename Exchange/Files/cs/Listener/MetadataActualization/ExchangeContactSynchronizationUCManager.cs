namespace Terrasoft.Configuration
{
	using System;
	using Terrasoft.Core;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeContactSynchronizationUCManager

	/// <summary>
	/// Class provides users that has enabled exchange calendars synchronization list.
	/// </summary>
	[DefaultBinding(typeof(ISynchronizationUCManager), Name = "ExchangeContact")]
	public class ExchangeContactSynchronizationUCManager : ExchangeSynchronizationUCManager
	{

		#region Properties: Public

		/// <summary>
		/// <see cref="ISynchronizationUCManager.SynchronizationControllerName"/>
		/// </summary>
		public override string SynchronizationControllerName => "ExchangeContactMetaDataActualizer";

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns related to <paramref name="entity"/> parent column name.
		/// </summary>
		/// <param name="entity"><see cref="Entity"/> instance.</param>
		/// <returns>Related to <paramref name="entity"/> activity unique identifier.</returns>
		protected override string GetRelatedEntityColumnName(Entity entity) {
			switch (entity.SchemaName) {
				case "Contact":
					return entity.Schema.GetPrimaryColumnName();
				case "ContactCommunication":
				case "ContactAddress":
					return "ContactId";
				default:
					return string.Empty;
			}
		}

		/// <summary>
		/// Searches users that can synchronize <paramref name="entityId"/> to exchange contacts list.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="entityId">Activity instance unique identifier.</param>
		/// <returns>Users that can synchronize <paramref name="entityId"/> to exchange contacts list.</returns>
		protected override Select GetUsersSelect(UserConnection userConnection, Guid entityId) {
			return new Select(userConnection)
					.Column("MSS", "CreatedById")
				.From("Contact").As("C")
				.InnerJoin("MailboxSyncSettings").As("MSS").On("C", "CreatedById").IsEqual("MSS", "CreatedById")
				.InnerJoin("ContactSyncSettings").As("CSS").On("CSS", "MailboxSyncSettingsId").IsEqual("MSS", "Id")
				.Where("C", "Id").IsEqual(Column.Parameter(entityId))
				.And("CSS", "ExportContacts").IsEqual(Column.Parameter(true))
				.And("MSS", "SynchronizationStopped").IsEqual(Column.Parameter(false)) as Select;
		}

		#endregion

	}

	#endregion

}
