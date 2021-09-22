namespace Terrasoft.Configuration.Exception {
	using System;

	#region Class: NotFoundPageEntityException

	/// <summary>
	/// The exception that is thrown when some steps of page defentition is failed.
	/// </summary>
	public class NotFoundPageEntityException : Exception {

		#region Constructors: Public

		public NotFoundPageEntityException(string message) : base(message) { }

		public NotFoundPageEntityException(string message, Exception ex) : base(message, ex) { }

		#endregion

	}

	#endregion

}
