namespace Terrasoft.Configuration.YanAnonymServices
{
    using System;
    using System.ServiceModel;
    using System.ServiceModel.Web;
    using System.ServiceModel.Activation;
    using Terrasoft.Core;
    using Terrasoft.Web.Common;
    using Terrasoft.Core.Entities; 

    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
    public class YanAnonymService: BaseService
    {
        // Ссылка на экземпляр UserConnection, требуемый для обращения к базе данных.
        private SystemUserConnection _systemUserConnection;
        private SystemUserConnection SystemUserConnection {
            get {
                return _systemUserConnection ?? (_systemUserConnection = (SystemUserConnection)AppConnection.SystemUserConnection);
            }
        }
        
        // Метод, возвращающий идентификатор контакта по имени контакта.
        [OperationContract]
        [WebInvoke(Method = "GET", RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped,
        ResponseFormat = WebMessageFormat.Json)]
        public string GetContactIdByName(string Name){
            // Указывается пользователь, от имени которого выполняется обработка данного http-запроса.
            SessionHelper.SpecifyWebOperationIdentity(HttpContextAccessor.GetInstance(), SystemUserConnection.CurrentUser);
            // Результат по умолчанию.
            var result = "";
            // Экземпляр EntitySchemaQuery, обращающийся в таблицу Contact базы данных.
            var esq = new EntitySchemaQuery(SystemUserConnection.EntitySchemaManager, "Contact");
            // Добавление колонок в запрос.
            var colId = esq.AddColumn("Id");
            var colName = esq.AddColumn("Name");
            // Фильтрация данных запроса.
            var esqFilter = esq.CreateFilterWithParameters(FilterComparisonType.Equal, "Name", Name);
            esq.Filters.Add(esqFilter);
            // Получение результата запроса.
            var entities = esq.GetEntityCollection(SystemUserConnection);
            // Если данные получены.
            if (entities.Count > 0)
            {
                // Возвратить значение колонки "Id" первой записи результата запроса.
                result = entities[0].GetColumnValue(colId.Name).ToString();
                // Также можно использовать такой вариант:
                // result = entities[0].GetTypedColumnValue<string>(colId.Name);
            }
            // Возвратить результат.
            return result;
        }
    }
}
