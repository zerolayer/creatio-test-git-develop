namespace Terrasoft.Sync.Exchange
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Common;
	using Terrasoft.Configuration;
	using Terrasoft.Core.DB;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.Sync;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	/// <summary>
	/// The base class object synchronization with the Exchange storage.
	/// </summary>
	public abstract class ExchangeBase : IRemoteItem
	{

		#region Constructors: Protected

		protected ExchangeBase(SyncItemSchema schema, Exchange.Item item, TimeZoneInfo timeZoneInfo) {
			Schema = schema;
			Item = item;
			if (item != null) {
				DateTime version = item.SafeGetValue<DateTime>(Exchange.ItemSchema.LastModifiedTime);
				Version = TimeZoneInfo.ConvertTimeFromUtc(version.ToUniversalTime(), timeZoneInfo);
			}
		}

		#endregion

		#region Properties: Protected

		protected string RemoteId = string.Empty;

		protected ExchangeUtilityImpl ExchangeUtility { get; } = new ExchangeUtilityImpl();

		#endregion

		#region Properties: Public

		/// <summary>
		/// Base item in Exchange storage.
		/// </summary>
		public Exchange.Item Item {
			get;
			protected set;
		}

		/// <summary>
		/// Unique item id in remote storage.
		/// </summary>
		public virtual string Id {
			get {
				if (string.IsNullOrEmpty(RemoteId)) {
					RemoteId = Item.Id.UniqueId;
				}
				return RemoteId;
			}
			internal set {
			}
		}

		/// <summary>
		/// Last modified date in external storage.
		/// </summary>
		public DateTime Version {
			get;
			private set;
		}

		/// <summary>
		/// Description mapping synchronization item.
		/// </summary>
		public SyncItemSchema Schema {
			get;
			private set;
		}

		/// <summary>
		/// Synchronization state element.
		/// </summary>
		public SyncState State {
			get;
			set;
		}

		/// <summary>
		/// Action on the synchronizing element.
		/// </summary>
		public SyncAction Action {
			get;
			set;
		}

		/// <summary>
		/// Display name.
		/// </summary>
		public virtual string DisplayName {
			get {
				return GetDisplayName();
			}
		}

		#endregion

		#region Methods: Protected

		protected static T GetEntityInstance<T>(SyncContext context, LocalItem localItem, string schemaName)
				where T : Terrasoft.Core.Entities.Entity {
			T instance;
			if (localItem.Entities[schemaName].Count == 0) {
				var schema = context.UserConnection.EntitySchemaManager.GetInstanceByName(schemaName);
				instance = (T)schema.CreateEntity(context.UserConnection);
				instance.SetDefColumnValues();
				localItem.AddOrReplace(schemaName, SyncEntity.CreateNew(instance));
			} else {
				var instanceSync = localItem.Entities[schemaName][0];
				instanceSync.Action = SyncAction.Update;
				instance = (T)instanceSync.Entity;
			}
			return instance;
		}

		protected bool IsDeletedProcessed(string schemaName, ref LocalItem localItem) {
			if (State == SyncState.Deleted) {
				if (localItem.Entities[schemaName].Count != 0) {
					localItem.Entities[schemaName][0].Action = SyncAction.Delete;
				}
				return true;
			}
			return false;
		}

		protected virtual string GetDisplayName() {
			return !string.IsNullOrEmpty(Item.Subject) ? Item.Subject : Id;
		}

		/// <summary>
		/// Sets irrelevant metadata sync local storage element tag "deleted".
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localId">Unique identifier item in local storage.</param>
		/// <param name="remoteId">Unique identifier synchronization item in remote storage.</param>
		protected void ActualizeOldMetadata(SyncContext context, Guid localId, String remoteId) {
			var select = new Select(context.UserConnection).Column("RemoteId").From("SysSyncMetaData")
				.Where("LocalId").IsEqual(Column.Parameter(localId))
				.And("SyncSchemaName").IsEqual(Column.Parameter("Activity"))
				.And("RemoteId").IsNotEqual(Column.Parameter(remoteId)) as Select;
			String prevRemoteId = string.Empty;
			using (var dbExecutor = context.UserConnection.EnsureDBConnection()) {
				using (var reader = select.ExecuteReader(dbExecutor)) {
					if (!reader.Read()) {
						return;
					}
					prevRemoteId = reader.GetColumnValue<string>("RemoteId");
				}
			}
			ItemMetadata prevMetaData = context.ReplicaMetadata.FindItemMetaData(prevRemoteId);
			if (prevMetaData.Count > 0) {
				foreach (Terrasoft.Core.Configuration.SysSyncMetaData metadata in prevMetaData) {
					metadata.LocalState = (int)SyncState.Deleted;
					metadata.RemoteState = (int)SyncState.Deleted;
					metadata.Save();
				}
			}
		}

		/// <summary>
		/// Set action for all synchronization schemes in local item.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="localItem">Sync element in local storage.</param>
		protected void SetLocalItemSchemasAction(SyncContext context, LocalItem entities, SyncAction action) {
			List<EntityConfig> schemas = entities.Schema.Configs;
			foreach (EntityConfig schema in schemas) {
				foreach (var syncEntity in entities.Entities[schema.SchemaName]) {
					context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
						"SetLocalItemSchemasAction action {0} for {1}", action, GetDisplayName());
					syncEntity.Action = action;
				}
			}
		}
		
		/// <summary>
		/// Checks, if the activity is old.
		/// </summary>
		/// <param name="dueDate">The due date of the activity.</param>
		/// <param name="context">Sync context.</param>
		/// <returns>Flag indicating if the appointment is old.</returns>
		protected bool IsOldActivity(DateTime dueDate, SyncContext context) {
			var syncProvider = (ExchangeActivitySyncProvider)context.RemoteProvider;
			var userSettings = syncProvider.GetUserSettings();
			var importActivitiesFrom = userSettings.ImportActivitiesFrom;
			return importActivitiesFrom > dueDate;
		}
		
		/// <summary>Returns activity instance for exchange appointment, in case of changed remote id.</summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="exchangeItem">Exchange item in external storage.</param>
		/// <param name="storedId">Id of bpm activity, stored in external property of exchange item.</param>
		/// <param name="localItem">Local storage item.</param>
		/// <returns>Activity instance.</returns>
		protected Entity GetSyncedActivityWithChangedRemoteId(SyncContext context, Exchange.Item exchangeItem, Guid storedId, LocalItem localItem) {
			Entity instance;
			var syncValueName = localItem.Schema.SyncValueName;
			var schema = context.UserConnection.EntitySchemaManager.GetInstanceByName("Activity");
			instance = schema.CreateEntity(context.UserConnection);
			if (!localItem.Entities["Activity"].Any(se => se.EntityId.Equals(storedId)) &&
					instance.FetchFromDB(storedId)) {
				SyncEntity syncEntity = SyncEntity.CreateNew(instance);
				var isCurrentUserOwner = context.UserConnection.CurrentUser.ContactId == instance.GetTypedColumnValue<Guid>("OwnerId");
				syncEntity.Action = isCurrentUserOwner ? SyncAction.Update : SyncAction.None;
				context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
					"GetSyncedActivityWithChangedRemoteId set action {0} for {1}",
					syncEntity.Action, GetDisplayName());
				localItem.AddOrReplace("Activity", syncEntity);
				if (syncValueName == ExchangeConsts.ExchangeAppointmentClassName) {
					context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
						"GetSyncedActivityWithChangedRemoteId ExchangeAppointmentClassName action update for {0}",
						GetDisplayName());
					Action = SyncAction.Update;
					syncEntity.Action = SyncAction.Update;
				} else {
					Action = isCurrentUserOwner ? SyncAction.Update : SyncAction.None;
					context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
						"GetSyncedActivityWithChangedRemoteId action {0} for {1}", Action, GetDisplayName());
					if (isCurrentUserOwner) {
						ActualizeOldMetadata(context, storedId, Id);
					}
				}
			} else {
						context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
					"GetSyncedActivityWithChangedRemoteId not found entity action {0} for {1}",
					Action, GetDisplayName());
				instance = GetEntityInstance<Entity>(context, localItem, "Activity");
			}
			return instance;
		}

		/// <summary>
		/// Returns activity instance.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="exchangeItem">Sync element in external storage.</param>
		/// <param name="localItem">Sync element in local storage.</param>
		protected Entity GetActivity(SyncContext context, Exchange.Item exchangeItem, ref LocalItem localItem) {
			Entity instance;
			Object localId;
			if (exchangeItem.TryGetProperty(ExchangeUtilityImpl.LocalIdProperty, out localId)) {
				context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
					"GetActivity use localId action {0} for {1}",
					Action, GetDisplayName());
				instance = GetSyncedActivityWithChangedRemoteId(context, exchangeItem, Guid.Parse(localId.ToString()), localItem);
			} else {
				context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload,
					"GetActivity not use localId action {0} for {1}",
					Action, GetDisplayName());
				instance = GetEntityInstance<Entity>(context, localItem, "Activity");
			}
			return instance;
		}

		/// <summary>
		/// Checks if remote entity already locked for synchronization.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <returns>
		/// <returns><c>True</c> if appointment already locked for synchronization, <c>false</c> otherwise.
		/// </returns>
		protected bool GetRemoteItemLockedForSync(SyncContext context) {
			return GetRemoteItemLockedForSync(context, Action);
		}

		/// <summary>
		/// Checks if remote entity already locked for synchronization.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		/// <param name="action"><see cref="SyncAction"/> instance.</param>
		/// <returns><c>True</c> if appointment already locked for synchronization, <c>false</c> otherwise.
		/// </returns>
		protected bool GetRemoteItemLockedForSync(SyncContext context, SyncAction action) {
			if (action != SyncAction.Create) {
				return false;
			}
			var helper = ClassFactory.Get<EntitySynchronizerHelper>();
			return !helper.CanCreateEntityInLocalStore(Id, context.UserConnection, "Exchange integration");
		}

		/// <summary>
		/// Checks if entity already locked for synchronization.
		/// </summary>
		/// <param name="entityId"><see cref="Entity"/> instance id.</param>
		/// <param name="context">Synchronization context.</param>
		/// <returns>
		/// True if entity already locked for synchronization, false otherwise.
		/// </returns>
		protected bool GetEntityLockedForSync(Guid entityId, SyncContext context) {
			if (Action != SyncAction.Create) {
				return false;
			}
			var helper = ClassFactory.Get<EntitySynchronizerHelper>();
			var result = !helper.CanCreateEntityInRemoteStore(entityId, context.UserConnection, "Exchange integration");
			if (result) {
				LogInfo(context, "Entity locked for sync in exchange (Id = {0}), sync action skipped.", entityId);
			}
			return result;
		}

		/// <summary>
		/// Writes information message to the log.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="operation"> Log action for an object synchronization.</param>
		/// <param name="direction">Synchronization direction.</param>
		/// <param name="format">Format.</param>
		/// <param name="args">Format patameters.</param>
		protected virtual void LogInfo(SyncContext context, SyncAction operation, SyncDirection direction,
				string format, params object[] args) {
			context.LogInfo(operation, direction, format, args);
		}

		/// <summary>
		/// Writes information message to the log.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="format">Format.</param>
		/// <param name="args">Format patameters.</param>
		protected virtual void LogInfo(SyncContext context, string format, params object[] args) {
			LogInfo(context, SyncAction.None, SyncDirection.DownloadAndUpload, format, args);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Fills BPM synchronization item<paramref name="localItem"/> 
		/// the values of the elements in the external storage.
		/// </summary>
		/// <param name="localItem">BPM item.</param>
		/// <param name="context">Synchronization context.</param>
		public abstract void FillLocalItem(SyncContext context, ref LocalItem localItem);

		/// <summary>
		/// Fills item in remote storage the values
		/// of the BPM element<paramref name="localItem"/>.
		/// </summary>
		/// <param name="localItem">BPM item.</param>
		/// <param name="context">Synchronization context.</param>
		public abstract void FillRemoteItem(SyncContext context, LocalItem localItem);

		#endregion
	}
}