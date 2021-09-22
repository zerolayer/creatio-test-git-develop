using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.IO;
using Terrasoft.Core;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{
	public class ITBSSshStorage : ITBSFileStorage
	{
		public ITBSSshStorage(UserConnection userConnection) : base(userConnection) { }

		public override BaseStorageSettings GetDefSettings(Entity storageEntity)
		{
			BaseStorageSettings settings = base.GetDefSettings(storageEntity);
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
				using (SftpClient sftp = new SftpClient(settings.Host, settings.Username, settings.Password))
				{
					sftp.Connect();
					Stream s = new MemoryStream(fileBytes);
					sftp.UploadFile(s, settings.PathToFile);
					sftp.Disconnect();
				}
				Dictionary<string, object> dicValues = new Dictionary<string, object>() {
					{ "ITBSFileLink", settings.PathToFile },
					{ "Size", fileBytes.Length }
				};
				result.Success = true;
				result.Values = dicValues;
			}
			catch (Exception)
			{
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "FtpSaveFileErrorMessage");
				result.Error = message;
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
				using (SftpClient sftp = new SftpClient(settings.Host, settings.Username, settings.Password))
				{
					sftp.Connect();
					using (MemoryStream fileStream = new MemoryStream())
					{
						sftp.DownloadFile(fileLink, fileStream);
						sftp.Disconnect();
						return fileStream.ToArray();
					}
				}
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
				using (SftpClient sftp = new SftpClient(settings.Host, settings.Username, settings.Password))
				{
					sftp.Connect();
					sftp.Disconnect();
					result.Success = true;
				}
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