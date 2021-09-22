namespace Terrasoft.Configuration.Users
{
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.AsyncOperations;
	using Terrasoft.Core.Entities.AsyncOperations.Interfaces;
	using Terrasoft.Core.Entities.Events;
	using Terrasoft.Core.Factories;

	#region Class: SysModuleGeneralEventListener

	/// <summary>
	/// Class provides SysUserInRole and SysAdminUnitInWorkplace entities events handling.
	/// </summary>
	[EntityEventListener(SchemaName = "SysUserInRole")]
	[EntityEventListener(SchemaName = "SysAdminUnitInWorkplace")]
	public class SysModuleGeneralEventListener : BaseEntityEventListener
	{

		#region Methods: Private

		/// <summary>
		/// Calls workplaces rights changed async operation.
		/// </summary>
		/// <param name="entity">Rights detail changed entity.</param>
		/// <param name="e">Entity change event arguments.</param>
		private void OnWorkplaceRightsChanged(Entity entity, EntityAfterEventArgs e) {
			var userConnection = entity.UserConnection;
			var asyncExecutor = ClassFactory.Get<IEntityEventAsyncExecutor>(new ConstructorArgument("userConnection", userConnection));
			var operationArgs = new EntityEventAsyncOperationArgs(entity, e);
			asyncExecutor.ExecuteAsync<UserInRoleEventAsyncOperation>(operationArgs);
		}

		#endregion

		/// <inhertidoc cref="BaseEntityEventListener.OnInserted"/>
		public override void OnInserted(object sender, EntityAfterEventArgs e) {
			base.OnInserted(sender, e);
			OnWorkplaceRightsChanged((Entity)sender, e);
		}

		/// <inheritdoc cref="BaseEntityEventListener.OnDeleted"/>
		public override void OnDeleted(object sender, EntityAfterEventArgs e) {
			base.OnDeleted(sender, e);
			OnWorkplaceRightsChanged((Entity)sender, e);
		}

		
	}

	#endregion

}
