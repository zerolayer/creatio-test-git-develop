using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Terrasoft.Core;
using Terrasoft.Core.DB;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{
	public class ITBSStorageHelper
	{
		#region Fields: Private

		private readonly UserConnection _userConnection;

		#endregion

		#region Constructor

		public ITBSStorageHelper(UserConnection userConnection)
		{
			_userConnection = userConnection;
			
		}

		#endregion

		#region Methods: Private

		private Entity GetEntity(string entityName, Guid entityId)
		{
			var esq = new EntitySchemaQuery(_userConnection.EntitySchemaManager, entityName);
			esq.UseAdminRights = false;
			esq.PrimaryQueryColumn.IsAlwaysSelect = true;
			esq.AddAllSchemaColumns();
			Entity entity = (entityId != Guid.Empty) ? esq.GetEntity(_userConnection, entityId) : null;
			return entity;
		}

		private void SetStorageStatus(OperationResult result, Entity storageEntity)
		{
			if (result.Success)
			{
				storageEntity.SetColumnValue("ITBSStatusId", ITBSConstCs.Storage.Status.Success);
				storageEntity.Save(false);
			}
			else
			{
				storageEntity.SetColumnValue("ITBSStatusId", ITBSConstCs.Storage.Status.Error);
				storageEntity.SetColumnValue("ITBSErrorDescription", result.Error);
				storageEntity.Save(false);
			}
		}

		#endregion

		#region Methods: Public

		public ITBSFileStorage GetFactoryInstant(Entity storageEntity)
		{
			ITBSFileStorage storageInstant = null;
			if (storageEntity != null)
			{
				Guid storageTypeId = storageEntity.GetTypedColumnValue<Guid>("ITBSTypeId");
				if (storageTypeId == ITBSConstCs.Storage.Type.Local) storageInstant = new ITBSLocalStorage(_userConnection);
				else if (storageTypeId == ITBSConstCs.Storage.Type.Network) storageInstant = new ITBSNetworkStorage(_userConnection);
				else if (storageTypeId == ITBSConstCs.Storage.Type.FTP) storageInstant = new ITBSFtpStorage(_userConnection);
				else if (storageTypeId == ITBSConstCs.Storage.Type.SSH) storageInstant = new ITBSSshStorage(_userConnection);
				else storageInstant = null;
			}
			else storageInstant = new ITBSDatabaseStorage(_userConnection);
			return storageInstant;
		}
		
		public OperationResult SaveFile(ITFileData data)
		{
			Entity storageEntity = GetEntity("ITBSFileStorageServer", data.StorageId);
			CheckAccess(new ITFileData() { StorageId = data.StorageId });
			ITBSFileStorage storageInstant = GetFactoryInstant(storageEntity);
			return storageInstant.SaveFile(storageEntity, data);
		}

		public OperationResult UpdateFile(ITFileData data)
		{
			Entity storageEntity = GetEntity("ITBSFileStorageServer", data.StorageId);
			CheckAccess(new ITFileData() { StorageId = data.StorageId });
			ITBSFileStorage storageInstant = GetFactoryInstant(storageEntity);
			return storageInstant.UpdateFile(storageEntity, data);
		}

		public virtual OperationResult CheckAccess(ITFileData data)
		{
			Entity storageEntity = GetEntity("ITBSFileStorageServer", data.StorageId);
			ITBSFileStorage storageInstant = GetFactoryInstant(storageEntity);
			OperationResult result = storageInstant.CheckAccess(storageEntity);
			if (storageEntity != null) SetStorageStatus(result, storageEntity);
			return result;
		}

		public ITReportData GetFile(Guid fileId)
		{
			Entity fileEntity = GetEntity("ITBSFile", fileId);
			Guid storageId = fileEntity.GetTypedColumnValue<Guid>("ITBSServerId");
			string fileName = fileEntity.GetTypedColumnValue<string>("Name");
			Entity storageEntity = GetEntity("ITBSFileStorageServer", storageId);
			ITBSFileStorage storageInstant = GetFactoryInstant(storageEntity);
			CheckAccess(new ITFileData() { StorageId = storageId });
			ITReportData fileData = new ITReportData()
			{
				Caption = System.IO.Path.GetFileNameWithoutExtension(fileName),
				Format = System.IO.Path.GetExtension(fileName).Replace(".", ""),
				Data = storageInstant.GetFile(storageEntity, fileEntity)
			};
			return fileData;
		}

		#endregion

	}
}

