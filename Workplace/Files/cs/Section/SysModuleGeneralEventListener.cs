namespace Terrasoft.Configuration.Section {
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Entities.AsyncOperations;
	using Terrasoft.Core.Entities.AsyncOperations.Interfaces;
	using Terrasoft.Core.Entities.Events;
	using Terrasoft.Core.Factories;
	using Terrasoft.Core.Store;

	#region Class: SysModuleGeneralEventListener

	/// <summary>
	/// Class provides SysModule entity events handling.
	/// </summary>
	[EntityEventListener(SchemaName = "SysModule")]
	public class SysModuleGeneralEventListener : BaseEntityEventListener {

		#region Methods: Protected

		/// <summary>
		/// Clears sections repository <see cref="ICacheStore"/> instance.
		/// </summary>
		/// <param name="sysModule"><see cref="Entity"/> instance.</param>
		/// <param name="e"><paramref name="sysModule"/> event arguments instance.</param>
		protected void ClearSectionRepositoryCache(Entity sysModule, EntityAfterEventArgs e) {
			var userConnection = sysModule.UserConnection;
			if (!sysModule.GetIsColumnValueLoaded("Type") || sysModule.GetTypedColumnValue<int>("Type") != (int)SectionType.General) {
				return;
			}
			var asyncExecutor = ClassFactory.Get<IEntityEventAsyncExecutor>(new ConstructorArgument("userConnection", userConnection));
			var operationArgs = new EntityEventAsyncOperationArgs(sysModule, e);
			asyncExecutor.ExecuteAsync<SysModuleGeneralEventAsyncOperation>(operationArgs);
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="BaseEntityEventListener.OnInserted"/>
		public override void OnInserted(object sender, EntityAfterEventArgs e) {
			base.OnInserted(sender, e);
			ClearSectionRepositoryCache((Entity)sender, e);
		}

		#endregion

	}

	#endregion

}
