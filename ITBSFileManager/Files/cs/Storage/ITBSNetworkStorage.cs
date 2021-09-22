using System;
using Terrasoft.Core;
using Terrasoft.Core.Entities;
using SharpCifs.Smb;
using System.Collections.Generic;
using System.IO;

namespace Terrasoft.Configuration
{
	public class ITBSNetworkStorage : ITBSFileStorage
	{
		public ITBSNetworkStorage(UserConnection userConnection) : base(userConnection) { }

		public override BaseStorageSettings GetDefSettings(Entity storageEntity)
		{
			BaseStorageSettings settings = base.GetDefSettings(storageEntity);
			settings.Prefix = "smb";
			settings.Separator = "/";
			return settings;
		}

		public override OperationResult SaveFile(Entity storageEntity, ITFileData fileData)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				string version = (fileData.File.Version != 1) ? "_" + fileData.File.Version.ToString() : String.Empty;
				string fileName = $"{fileData.File.FileContent.Caption}{version}.{fileData.File.FileContent.Format}";
				byte[] fileBytes = fileData.File.FileContent.Data;
				BaseStorageSettings settings = GetDefSettings(storageEntity);
				settings.FileName = fileName;
				var file = new SmbFile(settings.FullPathToFile, new NtlmPasswordAuthentication(settings.Domain, settings.Username, settings.Password));
				file.CreateNewFile();
				using (var writeStream = file.GetOutputStream())
				{
					writeStream.Write(fileBytes);
				}
				Dictionary<string, object> dicValues = new Dictionary<string, object>() {
					{ "ITBSFileLink", settings.FullPathToFile },
					{ "Size", fileBytes.Length }
				};
				result.Success = true;
				result.Values = dicValues;
			}
			catch (Exception ex)
			{
				result.Error = ex.Message;
			}
			return result;
		}

		public override string PerformSaveFileFS(Entity storageEntity, string path, string fileName, byte[] fileBytes)
		{
			throw new NotImplementedException();
		}

		public override OperationResult UpdateFile(Entity storageEntity, ITFileData data)
		{
			throw new NotImplementedException();
		}

		public override byte[] GetFile(Entity storageEntity, Entity fileEntity)
		{
			try
			{
				string fileLink = fileEntity.GetTypedColumnValue<string>("ITBSFileLink");
				BaseStorageSettings settings = GetDefSettings(storageEntity);
				var file = new SmbFile(fileLink, new NtlmPasswordAuthentication(settings.Domain, settings.Username, settings.Password));
				using (var readStream = file.GetInputStream())
				{
					var memStream = new MemoryStream();
					((Stream)readStream).CopyTo(memStream);
					return memStream.ToArray();
				};
			}
			catch (Exception) { }
			return null;
		}

		public override OperationResult DeleteFile(Entity storageEntity, Entity fileEntity)
		{
			throw new NotImplementedException();
		}

		public override OperationResult CheckAccess(Entity storageEntity)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				BaseStorageSettings settings = (BaseStorageSettings)GetDefSettings(storageEntity);
				var folder = new SmbFile(settings.FullPath, new NtlmPasswordAuthentication(settings.Domain, settings.Username, settings.Password));
				bool canRead = folder.CanRead();
				result.Success = true;
			}
			catch (Exception ex)
			{
				result.Error = ex.Message;
			}
			return result;
		}

		public override void CreateDirectory(string path)
		{
			throw new NotImplementedException();
		}

		public override bool ExistDirectory(string path)
		{
			throw new NotImplementedException();
		}

		public override bool ExistFile(string path)
		{
			throw new NotImplementedException();
		}

	}
}