using AlexPilotti.FTPS.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Terrasoft.Core;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{

	public class ITBSFtpStorage : ITBSFileStorage
	{
		public ITBSFtpStorage(UserConnection userConnection) : base(userConnection) { }

		public override BaseStorageSettings GetDefSettings(Entity storageEntity)
		{
			BaseStorageSettings settings = base.GetDefSettings(storageEntity);
			settings.Prefix = "ftp";
			settings.Separator = "/";
			return settings;
		}

		private FtpWebRequest GetFtpRequest(BaseStorageSettings settings, string path, string Method, byte[] fileBytes = null)
		{
			FtpWebRequest request = (FtpWebRequest)WebRequest.Create(path);
			request.Method = Method;
			request.Credentials = new NetworkCredential(settings.Username, settings.Password);
			request.EnableSsl = settings.UseSSL;
			request.UsePassive = settings.UsePassive;
			return request;
		}

		public void UploadFileToFtp(BaseStorageSettings settings, string fileName, byte[] fileBytes)
		{
			using (FTPSClient client = new FTPSClient())
			{
				ESSLSupportMode supportMode = (settings.UseSSL) ? ESSLSupportMode.Implicit : ESSLSupportMode.ClearText;
				client.Connect(settings.Host, settings.Port, new NetworkCredential(settings.Username, settings.Password), supportMode, null, null, 0, 0, 0, settings.Timeout);
				client.TryMakeDir(settings.Path);
				if (client.SetCurrentDirectory(settings.Path))
				{
					client.PutFile(fileName, fileBytes, fileName, null);
				}
			}
		}

		public void DeleteFileFromFtp(BaseStorageSettings settings, string pathToFile)
		{
			using (FTPSClient client = new FTPSClient())
			{
				ESSLSupportMode supportMode = (settings.UseSSL) ? ESSLSupportMode.Implicit : ESSLSupportMode.ClearText;
				client.Connect(settings.Host, settings.Port, new NetworkCredential(settings.Username, settings.Password), supportMode, null, null, 0, 0, 0, settings.Timeout);
				client.DeleteFile(pathToFile);
			}
		}

		public void UpdateFileToFtp(BaseStorageSettings settings, string fileName, byte[] fileBytes, string filePath)
		{
			using (FTPSClient client = new FTPSClient())
			{
				ESSLSupportMode supportMode = (settings.UseSSL) ? ESSLSupportMode.Implicit : ESSLSupportMode.ClearText;
				client.Connect(settings.Host, settings.Port, new NetworkCredential(settings.Username, settings.Password), supportMode, null, null, 0, 0, 0, settings.Timeout);
				client.PutFile(fileName, fileBytes, filePath, null);
			}
		}

		public override OperationResult SaveFile(Entity storageEntity, ITFileData fileData)
		{
			return SaveFileFS(storageEntity, fileData);
		}

		public override string PerformSaveFileFS(Entity storageEntity, string path, string fileName, byte[] fileBytes)
		{
			BaseStorageSettings settings = GetDefSettings(storageEntity);
			settings.AdditionalPath = path.Replace("\\","/");
			settings.FileName = fileName;
			UploadFileToFtp(settings, fileName, fileBytes);
			return settings.PathToFile;
		}
		
		public override OperationResult UpdateFile(Entity storageEntity, ITFileData data)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				BaseStorageSettings settings = GetDefSettings(storageEntity);
				UpdateFileToFtp(settings, data.File.FileContent.FileName, data.File.FileContent.Data, data.File.FileLink);
				Dictionary<string, object> dicValues = new Dictionary<string, object>();
				dicValues.Add("Size", data.File.FileContent.Data.Length);
				result.Success = true;
				result.Values = dicValues;
			}
			catch
			{
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "LocalSaveFileErrorMessage");
				result.Error = message;
			}
			return result;
		}

		public override byte[] GetFile(Entity storageEntity, Entity fileEntity)
		{
			try
			{
				string fileLink = fileEntity.GetTypedColumnValue<string>("ITBSFileLink");
				BaseStorageSettings settings = GetDefSettings(storageEntity);
				using (FTPSClient client = new FTPSClient())
				{
					ESSLSupportMode supportMode = (settings.UseSSL) ? ESSLSupportMode.Implicit : ESSLSupportMode.ClearText;
					client.Connect(settings.Host, settings.Port, new NetworkCredential(settings.Username, settings.Password), supportMode, null, null, 0, 0, 0, settings.Timeout);
					return client.GetFile(fileLink, true);
				}
			}
			catch{ }
			return null;
		}

		public override OperationResult DeleteFile(Entity storageEntity, Entity fileEntity)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				string fileLink = fileEntity.GetTypedColumnValue<string>("ITBSFileLink");
				BaseStorageSettings settings = GetDefSettings(storageEntity);
				DeleteFileFromFtp(settings, fileLink);
				result.Success = true;
			}
			catch
			{
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "LocalSaveFileErrorMessage");
				result.Error = message;
			}
			return result;
			
		}

		public override OperationResult CheckAccess(Entity storageEntity)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				BaseStorageSettings settings = (BaseStorageSettings)GetDefSettings(storageEntity);
				using (FTPSClient client = new FTPSClient())
				{
					ESSLSupportMode supportMode = (settings.UseSSL) ? ESSLSupportMode.Implicit : ESSLSupportMode.ClearText;
					client.Connect(settings.Host, settings.Port, new NetworkCredential(settings.Username, settings.Password), supportMode, null, null, 0, 0, 0, settings.Timeout);
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