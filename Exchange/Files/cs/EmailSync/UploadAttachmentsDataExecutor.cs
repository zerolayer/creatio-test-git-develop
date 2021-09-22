namespace Terrasoft.Configuration
{
	using System.Collections.Generic;
	using Terrasoft.Core;

	#region Class: UploadAttachmentsDataExecutor

	public class UploadAttachmentsDataExecutor : IJobExecutor
	{

		#region Methods: Public

		/// <summary>
		/// Executes exchange email attachments data synchronization.
		/// </summary>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="parameters">Synchronization synchronization (user email address etc.).</param>
		public virtual void Execute(UserConnection userConnection, IDictionary<string, object> parameters) {
			ExchangeUtility.UploadAttachmentsData(userConnection, parameters["UserEmailAddress"].ToString());
		}

		#endregion

	}

	#endregion

}

