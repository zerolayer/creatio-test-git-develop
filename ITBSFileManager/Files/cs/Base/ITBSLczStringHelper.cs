using Terrasoft.Common;
using Terrasoft.Core;

namespace Terrasoft.Configuration
{
	public class ITBSLczStringHelper
	{

		#region Fields: Private
		private readonly UserConnection _userConnection;
		#endregion

		#region Methods: Public
		public ITBSLczStringHelper(UserConnection userConnection)
		{
			_userConnection = userConnection;
		}

		/// <summary>
		/// Получить значение локализируемой строки
		/// </summary>
		/// <param name="moduleName"></param>
		/// <param name="lczName"></param>
		/// <param name="userConnection"></param>
		/// <returns></returns>
		public static string GetLczStringValue(UserConnection userConnection, string moduleName, string lczName)
		{
			string localizableStringName = string.Format("LocalizableStrings.{0}.Value", lczName);
			var localizableString = new LocalizableString(
				userConnection.Workspace.ResourceStorage, moduleName, localizableStringName);
			string value = localizableString.Value ??
							localizableString.GetCultureValue(GeneralResourceStorage.DefCulture, false);
			return value;
		}

		#endregion

	}
}