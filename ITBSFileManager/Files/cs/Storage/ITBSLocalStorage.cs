using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Security.AccessControl;
using Terrasoft.Common;
using Terrasoft.Core;
using Terrasoft.Core.DB;
using Terrasoft.Core.Entities;
using global::Common.Logging;

namespace Terrasoft.Configuration
{
	public class ITBSLocalStorage : ITBSFileStorage
	{
		ILog logger = LogManager.GetLogger("ITBSFileManager");

		public ITBSLocalStorage(UserConnection userConnection) : base(userConnection) { }

		private bool HasWritePermissionOnDir(string path)
		{
			var writeAllow = false;
			var writeDeny = false;
			var accessControlList = Directory.GetAccessControl(path);
			if (accessControlList == null)
				return false;
			var accessRules = accessControlList.GetAccessRules(true, true,
										typeof(System.Security.Principal.SecurityIdentifier));
			if (accessRules == null)
				return false;

			foreach (FileSystemAccessRule rule in accessRules)
			{
				if ((FileSystemRights.Write & rule.FileSystemRights) != FileSystemRights.Write)
					continue;

				if (rule.AccessControlType == AccessControlType.Allow)
					writeAllow = true;
				else if (rule.AccessControlType == AccessControlType.Deny)
					writeDeny = true;
			}

			return writeAllow && !writeDeny;
		}

		//public override OperationResult SaveFile(Entity storageEntity, ITFileData fileData)
		//{
		//	OperationResult result = new OperationResult() { Success = false };
		//	Dictionary<string, object> dicValues = new Dictionary<string, object>();
		//	try
		//	{
		//		string fileName = fileData.File.FileContent.FileName;
		//		byte[] fileBytes = fileData.File.FileContent.Data;
		//		string path = storageEntity.GetTypedColumnValue<string>("ITBSUrl");
		//		string fileNamePath = Path.Combine(path, fileName);
		//		if (!Directory.Exists(path)) CreateDirectory(path);
		//		System.IO.File.WriteAllBytes(fileNamePath, fileBytes);
		//		dicValues.Add("ITBSFileLink", fileNamePath);
		//		dicValues.Add("Size", fileBytes.Length);
		//		result.Values = dicValues;
		//		result.Success = true;
		//	}
		//	catch (Exception ex)
		//	{
		//		logger.Debug(ex.Message);
		//		var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "LocalSaveFileErrorMessage");
		//		result.Error = message;
		//	}
		//	return result;
		//}

		public override OperationResult SaveFile(Entity storageEntity, ITFileData fileData)
		{
			return SaveFileFS(storageEntity, fileData);
		}

		public override string PerformSaveFileFS(Entity storageEntity, string path, string fileName, byte[] fileBytes)
		{
			string fileNamePath = Path.Combine(path, fileName);
			if (!Directory.Exists(path)) CreateDirectory(path);
			System.IO.File.WriteAllBytes(fileNamePath, fileBytes);
			return fileNamePath;
		}

		public override OperationResult UpdateFile(Entity storageEntity, ITFileData data)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				string fileNamePath = data.File.FileLink;
				System.IO.File.WriteAllBytes(fileNamePath, data.File.FileContent.Data);
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
			string link = fileEntity.GetTypedColumnValue<string>("ITBSFileLink");
			return System.IO.File.ReadAllBytes(link);
		}

		public override OperationResult DeleteFile(Entity storageEntity, Entity fileEntity)
		{
			throw new NotImplementedException();
		}

		public override OperationResult CheckAccess(Entity storageEntity)
		{
			OperationResult result = new OperationResult() { Success = false };
			string path = storageEntity.GetTypedColumnValue<string>("ITBSUrl");
			try
			{
				bool hasRightAccess = HasWritePermissionOnDir(path);
				result.Success = hasRightAccess;
			}
			catch(Exception ex)
			{
				result.Error = ex.Message;
			}
			return result;
		}

		public override void CreateDirectory(string path)
		{
			Directory.CreateDirectory(path);
		}

		public override bool ExistDirectory(string path)
		{
			throw new NotImplementedException();
		}

		public override bool ExistFile(string path)
		{
			return File.Exists(path);
		}

	}
}