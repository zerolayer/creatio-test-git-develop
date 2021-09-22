namespace IntegrationV2.Files.cs.Listener.Subscription
{
	using IntegrationApi.Interfaces;
	using System;
	using Terrasoft.Common;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.Events;
	using Terrasoft.Core.Factories;
	using Terrasoft.Messaging.Common;

	#region Class: MailboxSyncSettingsEventListener

	[EntityEventListener(SchemaName = "MailboxSyncSettings")]
	public class MailboxSyncSettingsEventListener : BaseEntityEventListener
	{

		#region Properties: Private

		private IMsgChannelManager _channelManager;
		private IMsgChannelManager ChannelManager {
			get {
				if (_channelManager != null) {
					return _channelManager;
				}
				if (MsgChannelManager.IsRunning) {
					_channelManager = MsgChannelManager.Instance;
				}
				return _channelManager;
			}
		}

		#endregion

		#region Methods: Private

		private void SendInfoToClient(string eventName, string messageBody, Guid userId, string bodyType) {
			try {
				var channel = ChannelManager.FindItemByUId(userId);
				if (channel == null) {
					return;
				}
				var simpleMessage = new SimpleMessage {
					Id = userId,
					Body = messageBody,
				};
				simpleMessage.Header.Sender = eventName;
				simpleMessage.Header.BodyTypeName = bodyType;
				channel.PostMessage(simpleMessage);
			}
			catch (Exception) {
			}
		}

		private void SendInfoToClient(string eventName, Entity entity, Exception e) {
			SendInfoToClient(eventName, $"Error on deleting: {e}", entity.UserConnection.CurrentUser.Id, "Error");
		}
		
		private void SendInfoToClient(string eventName, Entity entity) {
			var messageBody = $"{{\"Id\":\"{entity.PrimaryColumnValue}\"," +
					$"\"Caption\": \"{entity.GetTypedColumnValue<string>("SenderEmailAddress")}\"," +
					$"\"IsShared\": \"{entity.GetTypedColumnValue<bool>("IsShared")}\" }}";
			SendInfoToClient(eventName, messageBody, entity.UserConnection.CurrentUser.Id, "Info");
		}

		#endregion

		#region Methods: Internal

		internal void SetMsgChannelManager(IMsgChannelManager channelManager) {
			_channelManager = channelManager;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnDeleting"/>
		/// </summary>
		public override void OnDeleting(object sender, EntityBeforeEventArgs e) {
			base.OnDeleting(sender, e);
			var entity = (Entity)sender;
			var userConnection = entity.UserConnection;
			var managerFactory = ClassFactory.Get<IListenerManagerFactory>();
			var manager = managerFactory.GetExchangeListenerManager(userConnection);
			try {
				manager.StopListener(entity.PrimaryColumnValue);
			} catch (Exception ex) {
				SendInfoToClient("SyncMsgLogger", entity, ex);
			}
		}

		/// <inheritdoc cref="BaseEntityEventListener.OnDeleted"/>
		public override void OnDeleted(object sender, EntityAfterEventArgs e) {
			base.OnDeleted(sender, e);
			var entity = (Entity)sender;
			SendInfoToClient("MailboxDeleted", entity);
		}

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnInserted"/>
		/// </summary>
		public override void OnInserted(object sender, EntityAfterEventArgs e) {
			base.OnInserted(sender, e);
			var entity = (Entity)sender;
			SendInfoToClient("MailboxAdded", entity);
		}

		/// <summary>
		/// <see cref="BaseEntityEventListener.OnUpdated"/>
		/// </summary>
		public override void OnUpdated(object sender, EntityAfterEventArgs e) {
			base.OnUpdated(sender, e);
			var entity = (Entity)sender;
			Guid oldTokenId = entity.GetTypedOldColumnValue<Guid>("OAuthTokenStorageId");
			Guid tokenId = entity.GetTypedColumnValue<Guid>("OAuthTokenStorageId");
			string oldPassword = entity.GetTypedOldColumnValue<string>("UserPassword");
			string password = entity.GetTypedColumnValue<string>("UserPassword");
			if ((oldPassword.IsNotNullOrEmpty() && password.IsNullOrEmpty())
					|| (oldTokenId.IsNotEmpty() && tokenId.IsEmpty())) {
				var userConnection = entity.UserConnection;
				var managerFactory = ClassFactory.Get<IListenerManagerFactory>();
				var manager = managerFactory.GetExchangeListenerManager(userConnection);
				try {
					manager.StopListener(entity.PrimaryColumnValue);
				} catch (Exception ex) {
					SendInfoToClient("SyncMsgLogger", entity, ex);
				}
			}
		}

		#endregion

	}

	#endregion

}
