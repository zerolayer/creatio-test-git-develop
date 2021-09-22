namespace Terrasoft.Configuration.Domain
{
	using System;
	using Terrasoft.Core.Factories;
	using TS = Terrasoft.Web.Http.Abstractions;

	#region Class: DomainResolver

	[DefaultBinding(typeof(IDomainResolver))]
	public class DomainFromRequestResolver : IDomainResolver
	{

		#region Fields: Private

		/// <summary>
		/// <see cref="HttpRequestWrapper"/> instance.
		/// </summary>
		private readonly TS.HttpRequest _request;

		#endregion

		#region Constructors: Public

		public DomainFromRequestResolver() {
			TS.HttpRequest currentRequest = TS.HttpContext.Current.Request ?? throw new Exception("Current request is empty.");
			_request = currentRequest;

		}

		#endregion

		private string GetUrl(TS.HttpRequest request, string path) {
			var combinedPath = string.Concat(request.ApplicationPath, path);
			var port = request.Host.Port ?? -1;
			var uriBuilder = new UriBuilder(request.Scheme, request.Host.Host, port, combinedPath);
			return uriBuilder.Uri.ToString().TrimEnd('/');
		}

		#region Methods: Public

		/// <inheritdoc cref="IDomainResolver.GetDomain"/>
		public string GetDomain() {
			return GetUrl(_request, string.Empty);
		}

		#endregion

	}

	#endregion

}