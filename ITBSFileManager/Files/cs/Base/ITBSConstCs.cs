using System;

namespace Terrasoft.Configuration
{

	public static class ITBSConstCs
	{

		public static class Storage
		{
			public static class Type
			{
				public static readonly Guid Local = new Guid("08D091DB-5486-4A12-A60E-9DD57AFC2569");
				public static readonly Guid Network = new Guid("75E91390-88D0-4A85-8F6A-4DF9B0C6D63F");
				public static readonly Guid FTP = new Guid("E444E00D-304F-46E8-8202-E8E702315BE3");
				public static readonly Guid ExtDataBase = new Guid("D12D8DF5-BEA0-4667-83E5-6AF003290E25");
				public static readonly Guid NFS = new Guid("A6AB4676-F346-478F-8CC1-E2F675353DCB");
				public static readonly Guid SSH = new Guid("E79DE7E9-96FC-4C6B-88BF-8A08BAE4FB43");
			}
			public static class Status
			{
				public static readonly Guid NotVerified = new Guid("DB361713-21D7-4585-B768-4A1EE688B375");
				public static readonly Guid Success = new Guid("F86B2D7F-6CA4-46E9-89A9-65CFAD29044F");
				public static readonly Guid Error = new Guid("50EA5B7E-F761-4FD9-AA14-812A7000D6F0");
			}
			public static class SaveMode
			{
				public static readonly Guid Classic = new Guid("D1EE9B60-0443-4EED-9DBE-06D5871E21F9");
				public static readonly Guid Recursively = new Guid("57C24273-A36E-4471-B563-8AC805B0384D");
			}
		}

	}

}
