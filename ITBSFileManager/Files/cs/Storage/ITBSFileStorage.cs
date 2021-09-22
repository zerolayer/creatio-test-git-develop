using System;
using System.Collections.Generic;
using System.IO;
using Terrasoft.Core;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{
	public class BaseStorageSettings
	{
		public string Host { get; set; }
		public string Username { get; set; }
		public string Domain { get; set; }
		public string Password { get; set; }
		public string Prefix { get; set; }
		public int Port { get; set; } = 21;
		public int Timeout { get; set; }
		public bool UseSSL { get; set; }
		public bool UsePassive { get; set; }
		public string DefaultPath { get; set; } = String.Empty;
		public string AdditionalPath { get; set; } = String.Empty;
		public string FileName { get; set; }
		public Guid StatusId { get; set; }
		public string Separator { get; set; }
		public string Path
		{
			get
			{
				string additionalPath = (AdditionalPath != String.Empty) ? $"{AdditionalPath}" : String.Empty;
				return $"{additionalPath}";
			}
		}
		public string PathToFile
		{
			get
			{
				return $"{Path}{Separator}{FileName}";
			}
		}
		public string FullPath
		{
			get
			{
				return $"{Prefix}:{new String(Separator[0], 2)}{Host}:{Port}{Separator}{Path}";
			}
		}
		public string FullPathToFile
		{
			get
			{
				return $"{FullPath}{Separator}{FileName}";
			}
		}
	}

	public abstract class ITBSFileStorage
	{
		public readonly UserConnection _userConnection;

		private static List<string> SplitString(string str)
		{
			List<string> list = new List<string>();
			int i = 0;
			while (i < str.Length - 1)
			{
				list.Add(str.Substring(i, 3));
				i += 3;
			}
			return list;
		}

		private string GetParentPath(string path)
		{
			int index = 0;
			while (path != String.Empty && path.Substring(path.Length - 3, 3).Contains("999"))
			{
				path = path.Substring(0, path.Length - 3);
				index += 1;
			}
			if (path != String.Empty) path = path.Substring(0, path.Length - 3) + (Convert.ToInt32(path.Substring(path.Length - 3, 3)) + 1).ToString().PadLeft(3, '0') + new string('0', index * 3);
			else path = new string('0', index * 3) + new string('0', 3);
			return path;
		}

		public virtual BaseStorageSettings GetDefSettings(Entity storageEntity)
		{
			string username = storageEntity.GetTypedColumnValue<string>("ITBSUserName");
			string password = storageEntity.GetTypedColumnValue<string>("ITBSPassword");
			string domain = storageEntity.GetTypedColumnValue<string>("ITBSDomain");
			string path = storageEntity.GetTypedColumnValue<string>("ITBSUrl").TrimStart('/').TrimEnd('/');
			string host = storageEntity.GetTypedColumnValue<string>("ITBSHost");
			int port = storageEntity.GetTypedColumnValue<int>("ITBSPort");
			bool useSSL = storageEntity.GetTypedColumnValue<bool>("ITBSUseSSL");
			Guid statusId = storageEntity.GetTypedColumnValue<Guid>("ITBSStatusId");
			int timeout = (int)Terrasoft.Core.Configuration.SysSettings.GetValue(_userConnection, "FtpConnectionTimeout");
			BaseStorageSettings settings = new BaseStorageSettings()
			{
				Host = host,
				Username = username,
				Domain = domain,
				Password = password,
				Port = port,
				StatusId = statusId,
				UsePassive = true,
				UseSSL = useSSL,
				Timeout = timeout,
				DefaultPath = path
			};
			return settings;
		}

		public OperationResult SaveFileFS(Entity storageEntity, ITFileData fileData)
		{
			string fileName = fileData.File.FileContent.FileName;
			byte[] fileBytes = fileData.File.FileContent.Data;
			string oldFilePathHash = storageEntity.GetTypedColumnValue<string>("ITBSFilePathHash");
			string filePathHash = oldFilePathHash;
			int fileNameHash = storageEntity.GetTypedColumnValue<int>("ITBSFileNameHash");
			char[] separates = new char[] { '\\', '/' };
			string startPath = storageEntity.GetTypedColumnValue<string>("ITBSUrl").TrimStart(separates).TrimEnd(separates);
			Dictionary<string, object> dicValues = new Dictionary<string, object>();
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				string fileExtention = Path.GetExtension(fileName);
				if (fileNameHash > 999)
				{
					fileNameHash = 0;
					filePathHash = GetParentPath(filePathHash);
				}
				string currentFilePathHash = filePathHash;
				currentFilePathHash = String.Join(@"\", SplitString(currentFilePathHash));
				string path = Path.Combine(startPath, currentFilePathHash);
				string newFileName = fileNameHash.ToString().PadLeft(3, '0') + fileExtention;
				
				string fileNamePath = PerformSaveFileFS(storageEntity, path, newFileName, fileBytes);

				fileNameHash += 1;
				storageEntity.SetColumnValue("ITBSFileNameHash", fileNameHash);
				storageEntity.SetColumnValue("ITBSFilePathHash", filePathHash);
				storageEntity.Save();
				dicValues.Add("ITBSFileLink", fileNamePath);
				dicValues.Add("Size", fileBytes.Length);
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

		public ITBSFileStorage(UserConnection userConnection)
		{
			_userConnection = userConnection;

		}

		public abstract OperationResult SaveFile(Entity storageEntity, ITFileData data);

		public abstract string PerformSaveFileFS(Entity storageEntity, string path, string fileNamePath, byte[] fileBytes);

		public abstract OperationResult UpdateFile(Entity storageEntity, ITFileData data);

		public abstract byte[] GetFile(Entity storageEntity, Entity fileEntity);

		public abstract OperationResult DeleteFile(Entity storageEntity, Entity fileEntity);

		public abstract OperationResult CheckAccess(Entity storageEntity);

		public abstract void CreateDirectory(string path);

		public abstract bool ExistDirectory(string path);

		public abstract bool ExistFile(string path);
	}
}