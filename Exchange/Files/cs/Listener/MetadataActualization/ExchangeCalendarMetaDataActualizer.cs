namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Terrasoft.Core.Factories;

	#region Class: ExchangeCalendarMetaDataActualizer

	/// <summary>
	/// Class implementation of <see cref="ISynchronizationController"/> for calendars metadata actualization.
	/// </summary>
	[DefaultBinding(typeof(ISynchronizationController), Name = "ExchangeCalendarMetaDataActualizer")]
	public class ExchangeCalendarMetaDataActualizer : SysSyncMetaDataActualizer
	{

		#region Methods: Protected

		/// <summary>
		/// Set additional parametrs for actualize metadatas.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected override void SetSyncParametrs(string entitySchemaName) {
			SchemaOrder = IsDetailSchemaName(entitySchemaName) ? 1 : 0;
			RemoteItemName = "ExchangeAppointment";
			StoreId = ExchangeConsts.AppointmentStoreId;
		}

		/// <summary>
		/// Get instance of <see cref="MetaDataInfo"/>.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected override MetaDataInfo GetMetaDataInfo(IDictionary<string, object> parameters) {
			return new ExchangeCalendarMetaDataInfo(parameters);
		}

		/// <summary>
		/// Indicates detals schema.
		/// </summary>
		/// <param name="syncSchemaName">Synchronization schema name.</param>
		/// <returns>True if schema is detail, otherwise false.</returns>
		protected override bool IsDetailSchemaName(string syncSchemaName) {
			return syncSchemaName != "Activity";
		}

		#endregion

	}

	#endregion

	#region Class: MetaDataParentInfo

	/// <summary>
	/// Class indicate meta datas of <see cref="ISynchronizationController"/> for calendars metadata actualization.
	/// </summary>
	public class ExchangeCalendarMetaDataInfo: MetaDataInfo
	{

		#region Constructor: Public

		public ExchangeCalendarMetaDataInfo(IDictionary<string, object> parameters) : base(parameters) {
		}

		#endregion

		#region Methods: Protected

		protected override void SetForeignColumnName() {
			switch (EntitySchemaName) {
				case "ActivityParticipant":
					ForeignColumnName = "ActivityId";
					break;
				case "Activity":
					ForeignColumnName = "Id";
					break;
				default:
					break;
			}
		}

		#endregion
	}

	#endregion

}