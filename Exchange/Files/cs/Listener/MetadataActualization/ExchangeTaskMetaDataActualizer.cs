namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Terrasoft.Core.Factories;

	#region Class: ContactMetaDataActualizer

	/// <summary>
	/// Class implementation of <see cref="ISynchronizationController"/> for tasks metadata actualization.
	/// </summary>
	[DefaultBinding(typeof(ISynchronizationController), Name = "ExchangeTaskMetaDataActualizer")]
	public class ExchangeTaskMetaDataActualizer : SysSyncMetaDataActualizer
	{

		#region Methods: Protected

		/// <summary>
		/// Set additional parametrs for actualize metadatas.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected override void SetSyncParametrs(string entitySchemaName) {
			SchemaOrder = IsDetailSchemaName(entitySchemaName) ? 1 : 0;
			RemoteItemName = "ExchangeTask";
			StoreId = ExchangeConsts.ExchangeTaskStoreId;
		}

		/// <summary>
		/// Get instance of <see cref="MetaDataInfo"/>.
		/// </summary>
		/// <param name="entitySchemaName"></param>
		protected override MetaDataInfo GetMetaDataInfo(IDictionary<string, object> parameters) {
			return new TaskMetaDataInfo(parameters);
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

	#region Class: ContactMetaDataInfo

	/// <summary>
	/// Class indicate meta datas of <see cref="ISynchronizationController"/> for tasks metadata actualization.
	/// </summary>
	public class TaskMetaDataInfo : MetaDataInfo
	{

		#region Constructor: Public

		public TaskMetaDataInfo(IDictionary<string, object> parameters) : base(parameters) {
		}

		#endregion

		#region Methods: Protected

		protected override void SetForeignColumnName() {
			switch (EntitySchemaName) {
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