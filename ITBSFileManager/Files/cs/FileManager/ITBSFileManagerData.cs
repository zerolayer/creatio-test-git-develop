using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Terrasoft.Configuration
{
	[Serializable]
	public class ITReportData
	{
		public string Caption { get; set; }
		public byte[] Data { get; set; }
		public string Format { get; set; }
		public string FileName {
			get { return Caption + "." + Format; }
		}
    }

	[DataContract]
	public class ITFile
	{
		public Guid Id { get; set; }
        public int Version { get; set; }
		public Guid ParentId { get; set; }
		public string FileLink { get; set; }
		public ITReportData FileContent { get; set; }
	}

	[DataContract]
	public class ITFileData
	{
		[DataMember]
		public Guid StorageId { get; set; }
		[DataMember]
		public Guid FolderId { get; set; }
        [DataMember]
        public string SchemaName { get; set; }
        [DataMember]
        public Guid RecordId { get; set; }
        [DataMember]
        public ITFile File { get; set; }
	}

	public class OperationResult
	{
		public bool Success { get; set; }
		public string Error { get; set; }
		public FileErrorType ErrorCode { get; set; }
		public Dictionary<string, object> Values { get; set; }
	}

	[DataContract]
	public enum SavingMode
	{
		None,
		Override,
		AsVersion
	}

	[DataContract]
	public enum FileErrorType
	{
		Success,
		FileExist = 100,
		AccessError = 200,
		BadRead = 201,
		BadPassword = 202,
		Other = 500
	}
}