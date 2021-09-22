namespace Terrasoft.Configuration.Workplace
{

	using System;
	using System.Collections.Generic;
	using System.Linq;
	using Terrasoft.Common;
	using Terrasoft.Configuration.Section;
	using Terrasoft.Core;
	using Terrasoft.Core.Factories;

	#region Class: WorkplaceManager

	[DefaultBinding(typeof(IWorkplaceManager))]
	public class WorkplaceManager : IWorkplaceManager
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="IWorkplaceRepository"/> implementation instance.
		/// </summary>
		private readonly IWorkplaceRepository _workplaceRepository;

		/// <summary>
		/// <see cref="ISectionRepository"/> implementation instance.
		/// </summary>
		private readonly ISectionRepository _sectionRepository;

		/// <summary>
		/// <see cref="IResourceStorage"/> implementation instance.
		/// </summary>
		private readonly IResourceStorage _resourceStorage;

		/// <summary>
		/// Current user identifier.
		/// </summary>
		private readonly Guid _currentUserId;

		#endregion

		#region Constructors: Public

		public WorkplaceManager(UserConnection uc) {
			_workplaceRepository = ClassFactory.Get<IWorkplaceRepository>(new ConstructorArgument("uc", uc));
			_sectionRepository = ClassFactory.Get<ISectionRepository>("General", new ConstructorArgument("uc", uc));
			_resourceStorage = uc.ResourceStorage;
			_currentUserId = uc.CurrentUser.Id;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates workplace exception message.
		/// </summary>
		/// <param name="exceptionMessageKey">Localizable string name.</param>
		/// <returns>Workplace exception message.</returns>
		private string GetWorkplaceExceptionMessage(string exceptionMessageKey) {
			return new LocalizableString(_resourceStorage, "SectionExceptionResources",
				$"LocalizableStrings.{exceptionMessageKey}.Value").ToString();
		}

		#endregion

		#region Methods: Protected

		/// <summary>
		/// Returns new workplace position.
		/// </summary>
		/// <param name="type">New workplace type.</param>
		/// <returns>New workplace position.</returns>
		protected int GetNewWorkplacePosition(WorkplaceType type) {
			var workplacePositions = GetWorkplacesByType(type).Select(w => w.Position);
			if (!workplacePositions.Any()) {
				return 0;
			}
			return workplacePositions.Max() + 1;
		}

		/// <summary>
		/// Retruns workplaces, which  position must be changed.
		/// </summary>
		/// <param name="workplace"><see cref="Workplace"/> that changed position instance.</param>
		/// <param name="position"><paramref name="workplace"/> new position.</param>
		/// <returns><see cref="Workplace"/> collection.</returns>
		protected IEnumerable<Workplace> GetWorkplacesToChange(Workplace workplace, int position) {
			var workplaces = GetWorkplacesByType(workplace.Type);
			IEnumerable<Workplace> result;
			if (workplace.Position > position) {
				result = workplaces.Where(w => w.Position >= position && w.Id != workplace.Id);
			} else {
				result = workplaces.Where(w => w.Position <= position && w.Id != workplace.Id);
			}
			return result.OrderBy(w => w.Position);
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc />
		public IEnumerable<Workplace> GetCurrentUserWorkplaces(Guid applicationClientTypeId) {
			return _workplaceRepository.GetAll().Where(w => w.ClientApplicationTypeId.Equals(applicationClientTypeId)
				&& w.GetIsAllowedForUser(_currentUserId));
		}

		/// <inheritdoc />
		public IEnumerable<Workplace> GetWorkplacesByType(WorkplaceType type) {
			var result = _workplaceRepository.GetAll().Where(w => w.Type == type);
			return result;
		}

		/// <inheritdoc />
		public Workplace CreateWorkplace(string name, WorkplaceType type) {
			var position = GetNewWorkplacePosition(type);
			var workplace = new Workplace(Guid.NewGuid(), name, type) {
				Position = position
			};
			if (_workplaceRepository.SaveWorkplace(workplace)) {
				return workplace;
			}
			var message = GetWorkplaceExceptionMessage("WorkplaceCreateFailed");
			throw new WorkplaceCreateFailedException(message);
		}

		/// <inheritdoc />
		public void ChangeName(Guid workplaceId, string name) {
			var workplace = _workplaceRepository.Get(workplaceId);
			workplace.SetName(name);
			_workplaceRepository.SaveWorkplace(workplace);
		}

		/// <inheritdoc />
		public void ChangePosition(Guid workplaceId, int position) {
			var workplace = _workplaceRepository.Get(workplaceId);
			if (workplace.Position == position) {
				return;
			}
			var workplacesToChange = GetWorkplacesToChange(workplace, position);
			var index = workplace.Position > position ? position + 1 : 0;
			foreach (var w in workplacesToChange) {
				if (w.Position != index) {
					w.Position = index;
					_workplaceRepository.SaveWorkplace(w);
				}
				index++;
			}
			workplace.Position = position;
			_workplaceRepository.SaveWorkplace(workplace);
		}

		/// <inheritdoc />
		public void DeleteWorkplace(Guid workplaceId) {
			_workplaceRepository.DeleteWorkplace(workplaceId);
		}

		/// <inheritdoc />
		public void AddSectionToWorkplace(Guid workplaceId, Guid sectionId) {
			var workplace = _workplaceRepository.Get(workplaceId);
			var section = _sectionRepository.Get(sectionId);
			workplace.AddSection(section.Id);
			_workplaceRepository.SaveWorkplace(workplace);
			_sectionRepository.ClearCache();
		}

		/// <inheritdoc />
		public void RemoveSectionFromWorkplace(Guid workplaceId, Guid sectionId) {
			var workplace = _workplaceRepository.Get(workplaceId);
			workplace.RemoveSection(sectionId);
			_workplaceRepository.SaveWorkplace(workplace);
			_sectionRepository.ClearCache();
		}

		/// <inheritdoc />
		public void ReloadWorkplaces() {
			_workplaceRepository.ClearCache();
		}

		#endregion

	}

	#endregion

}