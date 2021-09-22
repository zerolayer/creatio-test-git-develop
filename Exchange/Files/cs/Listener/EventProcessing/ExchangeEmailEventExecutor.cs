namespace Terrasoft.Configuration
{
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using IntegrationApi.Email;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Tasks;
	using Terrasoft.Sync;
	using Terrasoft.Sync.Exchange;

	#region Class: ExchangeEmailEventExecutor

	/// <summary>
	/// Class starts new exchange email synchronization.
	/// </summary>
	public class ExchangeEmailEventExecutor : BaseLoadEmailEventExecutor, IJobExecutor {

		#region Fields: Private

		/// <summary>
		/// Items identifier parameter name.
		/// </summary>
		private readonly string _itemIdsParameterName = "ItemIds";

		#endregion

		#region Methods: Private

		/// <summary>
		/// Select items identifiers from <paramref name="parameters"/>.
		/// </summary>
		/// <param name="parameters">Executor parameters collection.</param>
		/// <returns>Items identifiers.</returns>
		private IEnumerable<string> GetItemsIdsFromParameters(IDictionary<string, object> parameters) {
			return ((object[])parameters[_itemIdsParameterName]).Select(id => (string)id);
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Creates <see cref="RemoteProvider"/> instance.
		///  <see cref="ExchangeEmailEventsProvider"/> used as email remote provider.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Mailbox sender email address.</param>
		/// <returns><see cref="RemoteProvider"/> instance.</returns>
		protected RemoteProvider GetRemoteProvider(UserConnection userConnection, string senderEmailAddress) {
			return ClassFactory.Get<ExchangeEmailEventsProvider>(new ConstructorArgument("userConnection", userConnection),
				new ConstructorArgument("senderEmailAddress", senderEmailAddress), new ConstructorArgument("userSettings", null));
		}

		/// <summary>
		/// Creates <see cref="SyncContext"/> instance.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Executor parameters collection.</param>
		/// <returns><see cref="SyncContext"/> instance.</returns>
		protected SyncContext GetSyncContext(UserConnection userConnection, IDictionary<string, object> parameters) {
			var remoteProvider = GetRemoteProvider(userConnection, parameters["SenderEmailAddress"].ToString());
			var localProvider = new LocalProvider(userConnection);
			return new SyncContext {
				UserConnection = userConnection,
				LocalProvider = localProvider,
				RemoteProvider = remoteProvider,
				CurrentSyncStartVersion = DateTime.MinValue
			};
		}

		/// <summary>
		/// Creates and fills <see cref="LocalItem"/> instance.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="itemId">Email item identifier.</param>
		/// <returns><see cref="LocalItem"/> instance.</returns>
		protected LocalItem GetLocalItem(SyncContext context, string itemId) {
			var remoteProvider = context.RemoteProvider;
			context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload, "Start processing item {0}", itemId);
			SyncItemSchema schema = remoteProvider.SyncItemSchemaCollection.FirstOrDefault(s => s.SyncValueName == "ExchangeEmailMessage");
			IRemoteItem item = remoteProvider.LoadSyncItem(schema, itemId);
			LocalItem localItem = context.LocalProvider.FetchItem(null, schema, true);
			item.FillLocalItem(context, ref localItem);
			ProcessItemSyncAction(context, item);
			return localItem;
		}

		/// <summary>
		/// Additional actions after synchronization, based on paramref name="item"/> sync action.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="item">Current email <see cref="IRemoteItem"/> implementation instance.</param>
		protected void ProcessItemSyncAction(SyncContext context, IRemoteItem item) {
			if (item.Action == SyncAction.Repeat) {
				var itemId = item.Id;
				context.LogInfo(SyncAction.None, SyncDirection.DownloadAndUpload, "Item {0} synchronization need to be repeated", itemId);
				NeedReRun = true;
			}
		}

		/// <summary>
		/// Saves <paramref name="localItem"/> to database.
		/// </summary>
		/// <param name="context"><see cref="SyncContext"/> instance.</param>
		/// <param name="localItem"><see cref="LocalItem"/> instance.</param>
		/// <param name="itemId">External item identifier.</param>
		protected virtual void ApplyChanges(SyncContext context, LocalItem localItem, string itemId) {
			try {
				context.LocalProvider.ApplyChanges(context, localItem);
			} catch (Exception e) {
				context.LogError(SyncAction.Create, SyncDirection.DownloadAndUpload, "Error on item {0} processing", e, itemId);
			}
		}

		/// <summary>
		/// Commit changes.
		/// </summary>
		/// <param name="context">Synchronization context.</param>
		protected virtual void CommitChanges(SyncContext context) {
			context.RemoteProvider.CommitChanges(context);
		}

		/// <inheritdoc cref="BaseLoadEmailEventExecutor.Synchronize(UserConnection, IDictionary{string, object})"/>
		protected override void Synchronize(UserConnection uc, IDictionary<string, object> parameters) {
			var context = GetSyncContext(uc, parameters);
			foreach (var itemId in GetItemsIdsFromParameters(parameters)) {
				try {
					var localItem = GetLocalItem(context, itemId);
					ApplyChanges(context, localItem, itemId);
				} catch (Exception e) {
					context.LogError(SyncAction.None, SyncDirection.DownloadAndUpload,
						"Error on item {0} processing", e, itemId);
				}
			}
			CommitChanges(context);
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Loads and saves new exchange email.
		/// </summary>
		/// <param name="userConnection"><see cref="userConnection"/> instance.</param>
		/// <param name="parameters">Executor parameters collection.</param>
		public virtual void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			DoSyncAction(userConnection, parameters);
		}


		#endregion

	}

	#endregion

}
