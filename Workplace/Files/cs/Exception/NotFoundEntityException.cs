namespace Terrasoft.Configuration.Exception {
	using System;

	#region Class: NotFoundEntityException

	/// <summary>
	/// The exception that is thrown when entity does not found.
	/// </summary>
	public class NotFoundEntityException : Exception {

		#region Constructors: Public

		public NotFoundEntityException(string message) : base(message) { }

		public NotFoundEntityException(string message, Exception ex) : base(message, ex) { }

		#endregion

	}

	#endregion

}
