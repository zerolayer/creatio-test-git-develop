namespace Terrasoft.Configuration.Section {
	using Core;
	using Core.Entities.AsyncOperations;
	using Core.Entities.AsyncOperations.Interfaces;
	using Core.Factories;

	#region Class: SysModuleGeneralEventAsyncOperation

	/// <summary>
	/// Class implementats <see cref="IEntityEventAsyncOperation"/> interface for SysModule entity.
	/// </summary>
	public class SysModuleGeneralEventAsyncOperation : IEntityEventAsyncOperation {

		#region Methods: Private

		/// <summary>
		/// Creates <see cref="ISectionManager"/> implementation instance.
		/// </summary>
		/// <param name="type">Section manager type.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <returns><see cref="ISectionManager"/> implementation instance.</returns>
		private ISectionManager GetSectionManager(string type, UserConnection userConnection) {
			return ClassFactory.Get<ISectionManager>(new ConstructorArgument("uc", userConnection),
				new ConstructorArgument("sectionType", type));
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// <see cref="IEntityEventAsyncOperation.Execute"/>
		/// </summary>
		public void Execute(UserConnection userConnection, EntityEventAsyncOperationArgs arguments) {
			var manager = GetSectionManager("General", userConnection);
			manager.Save(arguments.EntityId);
		}

		#endregion

	}

	#endregion

}
