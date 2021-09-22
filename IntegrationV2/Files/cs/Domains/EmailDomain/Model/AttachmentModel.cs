namespace Terrasoft.EmailDomain.Model
{
	using System;

	#region Class: AttachmentModel

	/// <summary>
	/// Email attachment model.
	/// </summary>
	public class AttachmentModel
	{

		#region Properties: Public

		public string Name { get; set; }

		public byte[] Data { get; set; }

		public bool IsInline { get; set; }

		public Guid Id { get; set; }

		#endregion

	}

	#endregion

}
