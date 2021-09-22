using System.Collections.Generic;
using Terrasoft.Core;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{
	public class ITBSDatabaseStorage : ITBSFileStorage
	{
		public ITBSDatabaseStorage(UserConnection userConnection) : base(userConnection) { }

		public override OperationResult SaveFile(Entity storageEntity, ITFileData fileData)
		{
			return new OperationResult() { Success = true, Values = new Dictionary<string, object>() {
				{ "Data", fileData.File.FileContent.Data }
			}};
		}

		public override string PerformSaveFileFS(Entity storageEntity, string path, string fileName, byte[] fileBytes)
		{
			throw new System.NotImplementedException();
		}

		public override OperationResult UpdateFile(Entity storageEntity, ITFileData fileData)
		{
			return SaveFile(storageEntity, fileData);
		}

		public override byte[] GetFile(Entity storageEntity, Entity fileEntity)
		{
			return fileEntity.GetBytesValue("Data");
		}

		public override OperationResult DeleteFile(Entity storageEntity, Entity fileEntity)
		{
			throw new System.NotImplementedException();
		}

		public override OperationResult CheckAccess(Entity storageEntity)
		{
			return new OperationResult() { Success = true };
		}

		public override void CreateDirectory(string path)
		{
			throw new System.NotImplementedException();
		}

		public override bool ExistDirectory(string path)
		{
			throw new System.NotImplementedException();
		}

		public override bool ExistFile(string path)
		{
			throw new System.NotImplementedException();
		}

	}
}