using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Terrasoft.Common;
using Terrasoft.Core;
using Terrasoft.Core.DB;
using Terrasoft.Core.Entities;
using Terrasoft.Core.Factories;
using global::Common.Logging;

namespace Terrasoft.Configuration
{

	public class ITBSFileManagerHelper
	{
		#region Fields: Private

		private readonly UserConnection _userConnection;

		ILog logger = LogManager.GetLogger("ITBSFileManager");

		#endregion

		#region Constructor

		public ITBSFileManagerHelper(UserConnection userConnection)
		{
			_userConnection = userConnection;
		}

		#endregion

		#region Methods: Private

		private void InsertFileInFolder(ITFileData data)
		{
			ITBSEntityHelper.InsertEntity(_userConnection, "ITBSFileInFolder", new Dictionary<string, object>() {
				{ "ITBSFileId", data.File.Id },
				{ "ITBSFolderId", data.FolderId }
			});
		}

		private void InsertFile(ITFileData data, Dictionary<string, object> dicValues)
		{
			Dictionary<string, object> dic = new Dictionary<string, object>() {
				{ "Name", data.File.FileContent.FileName },
				{ "TypeId", new Guid("539BC2F8-0EE0-DF11-971B-001D60E938C6") },
				{ "Version", data.File.Version },
				{ "ITBSServerId", data.StorageId }
			};
			dic.AddRangeIfNotExists(dicValues);
			data.File.Id = ITBSEntityHelper.InsertEntity(_userConnection, "ITBSFile", dic);
		}

		private void UpdateFile(ITFileData data, Dictionary<string, object> dicValues)
		{
			ITBSEntityHelper.UpdateEntity(_userConnection, "ITBSFile", data.File.Id, dicValues);
		}

		private List<ITFile> GetSelectedFiles(List<Guid> files)
		{
			List<ITFile> list = new List<ITFile>();
			try
			{
				Select select = new Select(_userConnection)
					.Column("Id")
					.Column("Version")
					.Column("ITBSParentId")
					.From("ITBSFile")
					.Where("Id").In(Column.Parameters(files))
					as Select;
				using (DBExecutor dbExecutor = _userConnection.EnsureDBConnection())
				{
					using (IDataReader reader = select.ExecuteReader(dbExecutor))
					{
						while (reader.Read())
						{
							Guid fileId = reader.GetColumnValue<Guid>("Id");
							ITFile file = new ITFile()
							{
								Id = fileId,
								Version = reader.GetColumnValue<int>("Version"),
								ParentId = reader.GetColumnValue<Guid>("ITBSParentId"),
								FileContent = GetFileData(fileId)
							};
							list.Add(file);
						}
					}
				}
			}
			catch(Exception ex) {
				logger.Debug(ex.Message);
			}
			return list;
		}

		private byte[] CompressFiles(List<ITFile> files)
		{
			MemoryStream memoryStreamOfFile = new MemoryStream();
			ZipFile zip = new ZipFile();
			foreach (var file in files)
			{
				string suffix = (file.ParentId != Guid.Empty) ? "(" + file.Version.ToString() + ")" : String.Empty;
				string fileName = String.Format("{0}{1}.{2}", file.FileContent.Caption, suffix, file.FileContent.Format);
				zip.AddEntry(fileName, file.FileContent.Data);
			}
			zip.Save(memoryStreamOfFile);
			memoryStreamOfFile.Position = 0;
			return memoryStreamOfFile.ToArray();
		}

		private ITReportData GetCompressedFilesData(List<Guid> files, string archiveName = "archive")
		{
			List<ITFile> list = GetSelectedFiles(files);
			return new ITReportData()
			{
				Caption = archiveName,
				Format = "zip",
				Data = CompressFiles(list)
			};
		}

		private void UnzipFile(ITFileData fileData, ITFile file, string password)
		{
			Stream stream = new MemoryStream(file.FileContent.Data);
			ZipFile zip = ZipFile.Read(stream);
			zip.Password = password;
			foreach (ZipEntry e in zip)
			{
				using (MemoryStream memoryStreamOfFile = new MemoryStream())
				{
					e.Extract(memoryStreamOfFile);
					memoryStreamOfFile.Position = 0;
					byte[] data = memoryStreamOfFile.ToArray();
					ITFile extractFile = new ITFile()
					{
						Id = Guid.NewGuid(),
						ParentId = Guid.Empty,
						Version = 1,
						FileContent = new ITReportData()
						{
							Caption = Path.GetFileNameWithoutExtension(e.FileName),
							Format = Path.GetExtension(e.FileName).Replace(".", ""),
							Data = data
						}
					};
					fileData.File = extractFile;
					SaveFile(fileData, SavingMode.None);
				}
			}

		}

		private ITReportData GetFileData(Guid fileId)
		{
			var storageHelper = ClassFactory.Get<ITBSStorageHelper>();
			return storageHelper.GetFile(fileId);
		}

		private void UpdateFileParentId(List<Guid> files, Guid parentId)
		{
			try
			{
				Update insert = new Update(_userConnection, "ITBSFile")
					.Set("ITBSParentId", Column.Parameter(parentId))
					.Where("Id").In(Column.Parameters(files))
					as Update;
				insert.Execute();
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
			}
		}

		private void SetVersionParent(ITFileData data)
		{
			List<Guid> list = ITBSEntityHelper.ReadRecordsList<Guid>(_userConnection, "ITBSVwFileEntity", "Id", new Dictionary<string, object>() {
				{ "ITBSFolderId", data.FolderId },
				{ "ITBSFilename", data.File.FileContent.FileName },
				{ "ITBSServerId", data.StorageId }
			});
			if (list.Count > 0) UpdateFileParentId(list, data.File.Id);
		}

		private List<ITFile> GetFolderFiles(Guid folderId)
		{
			List<ITFile> list = new List<ITFile>();
			Select select = new Select(_userConnection)
				.Column("ITBSFileId")
				.Column("ITBSParentId")
				.Column("ITBSVersion")
				.From("ITBSVwFileEntity")
				.Where("ITBSFolderId").IsEqual(Column.Parameter(folderId))
				as Select;
			using (DBExecutor dbExecutor = _userConnection.EnsureDBConnection())
			{
				using (IDataReader reader = select.ExecuteReader(dbExecutor))
				{
					while (reader.Read())
					{
						list.Add(new ITFile()
						{
							Id = reader.GetColumnValue<Guid>("ITBSFileId"),
							ParentId = reader.GetColumnValue<Guid>("ITBSParentId"),
							Version = reader.GetColumnValue<int>("ITBSVersion")
						});
					}
				}
			}
			return list;
		}

		private void UpdateMoveFiles(List<ITFile> fileList, Guid destFolderId)
		{
			try
			{
				List<Guid> list = fileList.Select(x => x.Id).ToList();
				Update insert = new Update(_userConnection, "ITBSFileInFolder")
					.Set("ITBSFolderId", Column.Parameter(destFolderId))
					.Where("ITBSFileId").In(Column.Parameters(list))
					as Update;
				insert.Execute();
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
			}
		}

		private OperationResult ForceOverrideFile(ITFileData data)
		{
			OperationResult result = new OperationResult() { Success = false };
			var storageHelper = ClassFactory.Get<ITBSStorageHelper>();
			EntityCollection fileCollections = ITBSEntityHelper.ReadEntityCollection(_userConnection, "ITBSVwFileEntity", new string[] { "Id", "ITBSVersion", "ITBSFileLink" }, new Dictionary<string, object>() {
				{ "ITBSParent", null },
				{ "ITBSFilename", data.File.FileContent.FileName },
				{ "ITBSFolder", data.FolderId },
				{ "ITBSServer", data.StorageId }
			});
			if (fileCollections.Count > 0)
			{
				data.File.FileLink = fileCollections[0].GetTypedColumnValue<string>("ITBSFileLink");
				data.File.Version = fileCollections[0].GetTypedColumnValue<int>("ITBSVersion");
				data.File.Id = fileCollections[0].PrimaryColumnValue;
				result = storageHelper.UpdateFile(data);
				if (result.Success)
				{
					UpdateFile(data, result.Values);
				}
			}
			return result;
		}

		private OperationResult ForceVersionFile(ITFileData data)
		{
			OperationResult result = new OperationResult() { Success = false };
			var storageHelper = ClassFactory.Get<ITBSStorageHelper>();
			EntityCollection fileCollections = ITBSEntityHelper.ReadEntityCollection(_userConnection, "ITBSVwFileEntity", new string[] { "ITBSVersion", "ITBSFileLink" }, new Dictionary<string, object>() {
				{ "ITBSParent", null },
				{ "ITBSFilename", data.File.FileContent.FileName },
				{ "ITBSFolder", data.FolderId },
				{ "ITBSServer", data.StorageId }
			});
			result = storageHelper.SaveFile(data);
			if (result.Success)
			{
				data.File.Version = (fileCollections.Count > 0) ? fileCollections[0].GetTypedColumnValue<int>("ITBSVersion") + 1 : 1;
				InsertFile(data, result.Values);
				SetVersionParent(data);
				InsertFileInFolder(data);
			}
			return result;
		}

		#endregion

		#region Methods: Public

		/// <summary>
		/// Получение изображения
		/// </summary>
		/// <param name="fileId"></param>
		/// <returns></returns>
		public OperationResult GetImageFile(Guid fileId)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				ITReportData data = GetFileData(fileId);
                if (data?.Data != null)
                {
                    result.Values = new Dictionary<string, object>() {
                        {"FileKey", System.Convert.ToBase64String(data.Data) }
                    };
                    result.Success = true;
                }
                else
                {
                    var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OpenFileErrorMessage");
                    result.Error = message;
                }
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
				result.Error = message;
			}
			catch (Exception ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OpenFileErrorMessage");
				result.Error = message;
			}
			return result;
		}

		/// <summary>
		/// Получение файла
		/// </summary>
		/// <param name="fileId"></param>
		/// <param name="convertPdf"></param>
		/// <returns></returns>
		public OperationResult GetAppFile(Guid fileId, bool convertPdf)
		{
			OperationResult result = new OperationResult() { Success = false };
			string key = String.Empty;
			try
			{
				_userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
				ITReportData data = GetFileData(fileId);
                if (data?.Data != null)
                {
                    if (convertPdf)
                    {
                        data.Data = ITBSAsposePDFHelper.Convert(data.Data);
                        data.Format = "pdf";
                    }
                    key = string.Format("PreviewFileCacheKey_{0}", Guid.NewGuid());
                    _userConnection.SessionData[key] = data;
                    result.Values = new Dictionary<string, object>() {
                        {"FileKey", key }
                    };
                    result.Success = true;
                }
                else
                {
                    var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OpenFileErrorMessage");
                    result.Error = message;
                }
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
				result.Error = message;
			}
			catch (Exception ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OpenFileErrorMessage");
				result.Error = message;
			}
			return result;
		}

		/// <summary>
		/// Сохранение файла
		/// </summary>
		/// <param name="data"></param>
		/// <param name="mode"></param>
		/// <returns></returns>
		public OperationResult SaveFile(ITFileData data, SavingMode mode)
		{
			if (mode == SavingMode.None)
			{
				OperationResult checkResult = CheckExist(data);
				if (!checkResult.Success) return checkResult;
			}
			return ForceSaveFile(data, mode);
		}

		/// <summary>
		/// Принудительное сохранение файла
		/// </summary>
		/// <param name="data"></param>
		/// <returns></returns>
		public OperationResult ForceSaveFile(ITFileData data, SavingMode mode)
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				_userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
				if (mode == SavingMode.None || mode == SavingMode.AsVersion)
				{
					result = ForceVersionFile(data);
				}
				else if (mode == SavingMode.Override)
				{
					result = ForceOverrideFile(data);
				}
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				result.Error = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
			}
			return result;
		}

		/// <summary>
		/// Проверка на существание файла
		/// </summary>
		/// <param name="data"></param>
		/// <returns></returns>
		public OperationResult CheckExist(ITFileData data)
		{
			List<string> files = ITBSEntityHelper.ReadEntityCollection(_userConnection, "ITBSVwFileEntity", new string[] { "ITBSFile.Name" }, new Dictionary<string, object>() {
				{ "ITBSFolder", data.FolderId },
				{ "ITBSFilename", data.File.FileContent.FileName },
				{ "ITBSServer", data.StorageId }
			}).Select(x=>x.GetTypedColumnValue<string>("ITBSFile_Name")).ToList();
			return new OperationResult()
			{
				Success = (files.Count == 0),
				ErrorCode = (files.Count == 0) ? FileErrorType.Success : FileErrorType.FileExist,
				Error = (files.Count == 0) ? String.Empty : data.File.FileContent.FileName
			};
		}

		/// <summary>
		/// Проверка доступа к файлу
		/// </summary>
		/// <param name="data"></param>
		/// <returns></returns>
		public OperationResult CheckAccess(ITFileData data)
        {
            OperationResult result = new OperationResult() { Success = false };
            try
            {
                _userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
                var storageHelper = ClassFactory.Get<ITBSStorageHelper>();
                result = storageHelper.CheckAccess(data);
            }
            catch (LicException ex)
            {
				logger.Debug(ex.Message);
				result.Error = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
            }
            return result;
        }

		/// <summary>
		/// Скачивание файла
		/// </summary>
		/// <param name="files"></param>
		/// <returns></returns>
        public OperationResult DownloadFiles(List<Guid> files)
		{
			OperationResult result = new OperationResult() { Success = false };
			string key = String.Empty;
			try
			{
				_userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
				ITReportData data = (files.Count == 1) ? GetFileData(files[0]) : GetCompressedFilesData(files);
				if (data?.Data != null)
				{
					key = string.Format("ExportFileCacheKey_{0}", Guid.NewGuid());
					_userConnection.SessionData[key] = data;
					result.Values = new Dictionary<string, object>() {
						{"FileKey", key }
					};
					result.Success = true;
				}
				else
				{
					var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OpenFileErrorMessage");
					result.Error = message;
				}
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
				result.Error = message;
			}
			catch (FileNotFoundException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "FileNotFoundErrorMessage");
				result.Error = message;
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
			}
			return result;
		}

		/// <summary>
		/// Архивация файла
		/// </summary>
		/// <param name="files"></param>
		/// <param name="fileData"></param>
		/// <param name="archiveName"></param>
		/// <returns></returns>
		public OperationResult ZipFiles(List<Guid> files, ITFileData fileData, string archiveName)
		{
			OperationResult result = new OperationResult() { Success = false };
			Guid recordId = Guid.NewGuid();
			try
			{
				_userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
				ITReportData data = GetCompressedFilesData(files, archiveName);
				ITFile file = new ITFile()
				{
					Id = recordId,
					ParentId = Guid.Empty,
					Version = 1,
					FileContent = data
				};
				fileData.File = file;
				result = SaveFile(fileData, SavingMode.None);
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
				result.Error = message;
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
			}
			return result;
		}

		/// <summary>
		/// Распаковка файла
		/// </summary>
		/// <param name="files"></param>
		/// <param name="fileData"></param>
		/// <param name="password"></param>
		/// <returns></returns>
		public OperationResult UnzipFiles(List<Guid> files, ITFileData fileData, string password)
		{
			OperationResult result = new OperationResult() { Success = false };
			Guid recordId = Guid.NewGuid();
			string key = String.Empty;
			try
			{
				_userConnection.LicHelper.CheckHasOperationLicense("FileManagerFileOperation.Use");
				List<ITFile> list = GetSelectedFiles(files);
				foreach (ITFile item in list)
				{
					UnzipFile(fileData, item, password);
				}
				result.Success = true;
			}
			catch (LicException ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(_userConnection, "ITBSLocalizableHelper", "OperationWithoutLicErrorMessage");
				result.Error = message;
			}
			catch (BadReadException ex)
			{
				logger.Debug(ex.Message);
				result.ErrorCode = FileErrorType.BadRead;
			}
			catch (BadPasswordException ex)
			{
				logger.Debug(ex.Message);
				result.ErrorCode = FileErrorType.BadPassword;
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
				result.ErrorCode = FileErrorType.Other;
			}
			return result;
		}

		/// <summary>
		/// Перемещение файла
		/// </summary>
		/// <param name="files"></param>
		/// <param name="srcFolderId"></param>
		/// <param name="destFolderId"></param>
		/// <returns></returns>
		public OperationResult MoveFiles(List<Guid> files, Guid srcFolderId, Guid destFolderId)
		{
			OperationResult result = new OperationResult() { Success = false };
			string key = String.Empty;
			try
			{
				List<ITFile> list = GetFolderFiles(srcFolderId);
				list = list.Where(x => files.Contains(x.ParentId) || files.Contains(x.Id) && x.ParentId == Guid.Empty).ToList();
				UpdateMoveFiles(list, destFolderId);
			}
			catch (Exception ex) {
				logger.Debug(ex.Message);
			}
			return result;
		}

		#endregion
	}
}

