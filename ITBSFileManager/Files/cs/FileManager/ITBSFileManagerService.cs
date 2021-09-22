using System;
using System.Collections.Generic;
using System.IO;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Web;
using Terrasoft.Core;
using Terrasoft.Web.Common;
using global::Common.Logging;

namespace Terrasoft.Configuration
{
	[ServiceContract]
	[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
	public class ITBSFileManagerService : BaseService
	{
		ILog logger = LogManager.GetLogger("ITBSFileManager");

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult ImportFile()
		{
			OperationResult result = new OperationResult() { Success = false };
			try
			{
				var httpContext = HttpContextAccessor.GetInstance();
				var request = httpContext.Request;
				var queryString = request.Form;
				Guid folderId = new Guid(queryString["parentColumnValue"]);
				Guid storageId = new Guid(queryString["storageId"]);
				string parentName = queryString["parentName"];
                string schemaName = queryString["schemaName"];
                Guid recordId = new Guid(queryString["recordId"]);
                SavingMode mode = (SavingMode)Convert.ToInt32(queryString["mode"]);
				System.IO.Stream fileContent = request.Files["files"].InputStream;
				string fileName = request.Files["files"].FileName;
				fileContent.Position = 0;
				int FileSize = (int)fileContent.Length;
				byte[] fileData = new byte[fileContent.Length];
				fileContent.Read(fileData, 0, fileData.Length);
				ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
				string PrefileName = (parentName != String.Empty) ? parentName : fileName;
				ITFileData data = new ITFileData()
				{
					StorageId = storageId,
					FolderId = folderId,
                    SchemaName = schemaName,
                    RecordId = recordId,
					File = new ITFile()
					{
						FileContent = new ITReportData()
						{
							Caption = Path.GetFileNameWithoutExtension(PrefileName),
							Format = Path.GetExtension(PrefileName).Replace(".", ""),
							Data = fileData
						}
					}
				};
				result = helper.SaveFile(data, mode);
			}
			catch (Exception ex)
			{
				logger.Debug(ex.Message);
				var message = ITBSLczStringHelper.GetLczStringValue(UserConnection, "ITBSLocalizableHelper", "SaveFileErrorMessage");
				result.Error = message;
			}
			return result;
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult CheckExist(Guid storageId, SavingMode mode)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			ITFileData data = new ITFileData()
			{
				StorageId = storageId,
			};
			return helper.CheckAccess(data);
		}

		[OperationContract]
        [WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
        public OperationResult CheckAccess(Guid storageId)
        {
            ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
            ITFileData data = new ITFileData()
            {
                StorageId = storageId,
            };
            return helper.CheckAccess(data);
        }

        [OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult DownloadFiles(List<Guid> files)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.DownloadFiles(files);
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult ZipFiles(List<Guid> files, ITFileData fileData, string archiveName)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.ZipFiles(files, fileData, archiveName);
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult UnzipFiles(List<Guid> files, ITFileData fileData, string password)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.UnzipFiles(files, fileData, password);
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult MoveFiles(List<Guid> files, Guid srcFolderId, Guid destFolderId)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.MoveFiles(files, srcFolderId, destFolderId);
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult GetImageFile(Guid fileId)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.GetImageFile(fileId);
		}

		[OperationContract]
		[WebInvoke(Method = "POST", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
			ResponseFormat = WebMessageFormat.Json)]
		public OperationResult GetAppFile(Guid fileId, bool convertPdf)
		{
			ITBSFileManagerHelper helper = new ITBSFileManagerHelper(UserConnection);
			return helper.GetAppFile(fileId, convertPdf);
		}

		[OperationContract]
		[WebGet(UriTemplate = "GetFile/{key}")]
		public Stream GetFile(string key)
		{
			var reportObj = UserConnection.SessionData[key];
			UserConnection.SessionData.Remove(key);
			ITReportData file = reportObj as ITReportData;
			var fileCaption = $"{file.Caption}.{file.Format}";
			var outgoingResponse = WebOperationContext.Current.OutgoingResponse;
			outgoingResponse.ContentType = (file.Format == "csv")
				? "application/csv; charset=UTF-8"
				: "application/octet-stream; charset=UTF-8";
			var fileStream = new MemoryStream(file.Data);
			outgoingResponse.ContentLength = fileStream.Length;
			outgoingResponse.Headers.Add("Content-Disposition", "attachment; filename*=UTF-8''" + HttpUtility.UrlPathEncode(fileCaption) + "");
			return fileStream;
		}

		public string GetContentType(string format)
		{
			string contentType = String.Empty;
			switch (format)
			{
				case "pdf":
					contentType = "application/pdf";
					break;
				case "txt":
					contentType = "text/plain";
					break;
				default:
					break;
			}
			return contentType;
		}

		[OperationContract]
		[WebGet(UriTemplate = "GetPreviewFile/{key}/{media}")]
		public Stream GetPreviewFile(string key, string media)
		{
			var reportObj = UserConnection.SessionData[key];
			if (!Convert.ToBoolean(media)) UserConnection.SessionData.Remove(key);
			ITReportData file = reportObj as ITReportData;
			var fileCaption = $"{file.Caption}.{file.Format}";
			string contentType = GetContentType(file.Format);
			var outgoingResponse = WebOperationContext.Current.OutgoingResponse;
			outgoingResponse.ContentType = String.Format("{0}; charset=UTF-8", contentType);
			var fileStream = new MemoryStream(file.Data);
			outgoingResponse.ContentLength = fileStream.Length;
			outgoingResponse.Headers.Add("Content-Disposition", "filename*=UTF-8''" + HttpUtility.UrlPathEncode(fileCaption) + "");
			return fileStream;
		}
	}
}
