namespace Terrasoft.EmailDomain
{
	using System;
	using System.Collections.Generic;
	using EmailContract.DTO;
	using IntegrationApi.Interfaces;
	using Terrasoft.Core;
	using Terrasoft.Core.Entities;
	using Terrasoft.Core.Factories;
	using Terrasoft.EmailDomain.Interfaces;
	using Terrasoft.EmailDomain.Model;

	#region Class: AttachmentRepository

	/// <summary>
	/// Attachment model repository.
	/// </summary>
	[DefaultBinding(typeof(IAttachmentRepository))]
	internal class AttachmentRepository : IAttachmentRepository
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="UserConnection"/> instance.
		/// </summary>
		private readonly UserConnection _userConnection;

		#endregion

		#region Constructors: Public

		public AttachmentRepository(UserConnection uc) {
			_userConnection = uc;
		}

		#endregion

		#region Methods: Private

		/// <summary>
		/// Creates activity file entity.
		/// </summary>
		/// <param name="attach">Attachment model instance.</param>
		/// <param name="emailId">Email identifier.</param>
		/// <returns>Activity file entity.</returns>
		private Entity CreateActivityFileEntity(AttachmentModel attach, Guid emailId) {
			var activityFile = GetActivityFileEntity();
			var utils = ClassFactory.Get<IActivityUtils>();
			activityFile.SetDefColumnValues();
			var name = utils.GetAttachmentName(_userConnection, attach.Name);
			activityFile.PrimaryColumnValue = attach.Id;
			activityFile.SetColumnValue("Name", name);
			activityFile.SetColumnValue("Version", 1);
			activityFile.SetColumnValue("TypeId", IntegrationConsts.FileTypeId);
			activityFile.SetColumnValue("ActivityId", emailId);
			activityFile.SetColumnValue("ModifiedOn", DateTime.Now);
			activityFile.SetColumnValue("Uploaded", true);
			activityFile.SetColumnValue("Inline", attach.IsInline);
			return activityFile;
		}

		/// <summary>
		/// Saves <paramref name="attach"/> as <paramref name="emailId"/> activity file.
		/// </summary>
		/// <param name="attach">Attachment model instance.</param>
		/// <param name="emailId">Email identifier.</param>
		private void SaveActivityFile(AttachmentModel attach, Guid emailId) {
			var uploader = ClassFactory.Get<IFileUploader>("EmailAttachmentUploader", new ConstructorArgument("uc", _userConnection));
			var activityFile = CreateActivityFileEntity(attach, emailId);
			activityFile.Save();
			uploader.UploadAttach(activityFile.PrimaryColumnValue, activityFile.GetTypedColumnValue<string>("Name"), attach.Data);
		}

		/// <summary>
		/// Create empty activity file <see cref="Entity"/>.
		/// </summary>
		/// <returns>Activity file <see cref="Entity"/>.</returns>
		private Entity GetActivityFileEntity() {
			var schema = _userConnection.EntitySchemaManager.GetInstanceByName("ActivityFile");
			return schema.CreateEntity(_userConnection);
		}

		/// <summary>
		/// Get activity files <see cref="EntityCollection"/>.
		/// </summary>
		/// <param name="activityId"></param>
		/// <returns>ActivityFile Entity collection</returns>
		private EntityCollection GetActivityFileEntities(EntitySchemaQuery activityFileESQ, Guid activityId) {
			activityFileESQ.PrimaryQueryColumn.IsAlwaysSelect = true;
			activityFileESQ.AddColumn("Inline").Name = "Inline";
			activityFileESQ.Filters.Add(activityFileESQ.CreateFilterWithParameters(
				FilterComparisonType.Equal, "Activity", activityId));
			return activityFileESQ.GetEntityCollection(_userConnection);
		}

		#endregion

		#region Methods: Public

		/// <inheritdoc cref="IAttachmentRepository.GetAttachments(Guid)"/>
		public List<Attachment> GetAttachments(Guid activityId) {
			var attachLoader = ClassFactory.Get<IAttachmentFileLoader>(
				new ConstructorArgument("userConnection", _userConnection));
			var activityFileESQ = new EntitySchemaQuery(_userConnection.EntitySchemaManager, "ActivityFile");
			var activityFiles = GetActivityFileEntities(activityFileESQ, activityId);
			var result = new List<Attachment> { };
			foreach (var activityFile in activityFiles) {
				Guid id = activityFile.PrimaryColumnValue;
				var attachment = attachLoader.GetAttachment(activityFileESQ.RootSchema.UId, id);
				attachment.IsInline = activityFile.GetTypedColumnValue<bool>("Inline");
				result.Add(attachment);
			}
			return result;
		}

		/// <inheritdoc cref="IAttachmentRepository.GetAttachmentLink(Guid)"/>
		public string GetAttachmentLink(Guid attachmentId) {
			var item = _userConnection.EntitySchemaManager.GetItemByName("ActivityFile");
			return $"../rest/FileService/GetFile/{item.UId}/{attachmentId}";
		}

		/// <inheritdoc cref="IAttachmentRepository.SaveAttachments(EmailModel)"/>
		public void SaveAttachments(EmailModel email) {
			foreach (var attach in email.Attachments) {
				SaveActivityFile(attach, email.Id);
			}
		}

		/// <inheritdoc cref="IAttachmentRepository.SetInline(Guid)"/>
		public void SetInline(Guid attachmentId) {
			var activityFileEntity = GetActivityFileEntity();
			if (activityFileEntity.FetchFromDB(attachmentId, false)) {
				activityFileEntity.SetColumnValue("Inline", true);
				activityFileEntity.Save();
			}
		}

		#endregion

	}

	#endregion

}
