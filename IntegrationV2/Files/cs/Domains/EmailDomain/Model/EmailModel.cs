namespace Terrasoft.EmailDomain.Model
{
	using System;
	using System.Collections.Generic;
	using EmailContract;

	#region Class: EmailModel

	/// <summary>
	/// Email model.
	/// </summary>
	public class EmailModel : EmailModelHeader
	{

		#region Properties: Public

		public string Subject { get; set; }

		public string Body { get; set; }

		public string From { get; set; }

		public List<string> To { get; set; }

		public List<string> Copy { get; set; }

		public List<string> BlindCopy { get; set; }

		public DateTime SendDate { get; set; }

		public bool IsHtmlBody { get; set; }

		public Guid OwnerId { get; set; }

		public EmailImportance Importance { get; set; }

		internal string OriginalBody { get; set; }

		public List<string> Headers { get; set; }

		public List<AttachmentModel> Attachments { get; set; }  = new List<AttachmentModel>();

		#endregion

	}

	#endregion

}
