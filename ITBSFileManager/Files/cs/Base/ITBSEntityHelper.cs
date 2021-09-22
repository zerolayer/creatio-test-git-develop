using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Terrasoft.Common;
using Terrasoft.Core;
using Terrasoft.Core.DB;
using Terrasoft.Core.Entities;

namespace Terrasoft.Configuration
{
	public static class EsqExtension
	{
		public static void AddColumns(this EntitySchemaQuery esq, List<string> columns)
		{
			columns.ForEach(column => esq.AddColumn(column));
		}
	}
	
	public class ITBSEntityHelper
	{
		/// <summary>
		/// Получить идентификатор записи
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="schema"></param>
		/// <param name="credentials"></param>
		/// <returns></returns>
		public static Guid SelectEntityId(UserConnection userConnection, string schema, Dictionary<string, object> credentials)
		{
			Select select = new Select(userConnection)
				.Top(1)
				.Column("Id")
				.From(schema) as Select;
			foreach (KeyValuePair<string, object> key in credentials)
			{
				var indexOf = credentials.Keys.ToList().IndexOf(key.Key);
				if (indexOf == 0)
					select.Where(key.Key).IsEqual(Column.Parameter(key.Value));
				else
					select.And(key.Key).IsEqual(Column.Parameter(key.Value));
			}
			return select.ExecuteScalar<Guid>();
		}

		/// <summary>
		/// Обновить сущность
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="id"></param>
		/// <param name="fieldsNameFieldsValue"></param>
		public static void UpdateEntity(UserConnection userConnection, string entityName, Guid id,
			Dictionary<string, object> fieldsNameFieldsValue)
		{
			var result = string.Empty;
			var entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			var entity = entitySchema.CreateEntity(userConnection);
			if (entity.FetchFromDB(id, false))
			{
				try
				{
					foreach (var FildNameFildValue in fieldsNameFieldsValue)
					{
						if (FildNameFildValue.Value is string)
						{
							string fildNameFildValueValue = FildNameFildValue.Value.ToString();
							if (!string.IsNullOrEmpty(fildNameFildValueValue))
							{
								var textDataValueType =
									(TextDataValueType)
										entity.Schema.Columns.GetByColumnValueName(FildNameFildValue.Key).DataValueType;
								int size = textDataValueType.Size;
								if (!textDataValueType.IsMaxSize && (fildNameFildValueValue.Length > size))
								{
									fildNameFildValueValue = fildNameFildValueValue.Substring(0, size);
								}
								entity.SetColumnValue(FildNameFildValue.Key, fildNameFildValueValue);
							}
							else
							{
								entity.SetColumnValue(FildNameFildValue.Key, String.Empty);
							}
						}
						else if (FildNameFildValue.Value is DateTime)
						{
							if ((DateTime) FildNameFildValue.Value != DateTime.MinValue)
							{
								entity.SetColumnValue(FildNameFildValue.Key, FildNameFildValue.Value);
							}
						}
						else if (FildNameFildValue.Value is Guid)
						{
							if ((Guid) FildNameFildValue.Value != Guid.Empty)
							{
								entity.SetColumnValue(FildNameFildValue.Key, FildNameFildValue.Value);
							}
						}
						else if (FildNameFildValue.Value is byte[])
						{
							if ((byte[])FildNameFildValue.Value != new byte[0])
							{
								entity.SetBytesValue(FildNameFildValue.Key, (byte[])FildNameFildValue.Value);
							}
						}
						else if (FildNameFildValue.Value != null)
						{
							entity.SetColumnValue(FildNameFildValue.Key, FildNameFildValue.Value);
						}
						else
						{
							if (FildNameFildValue.Value == null)
							{
								entity.SetColumnValue(FildNameFildValue.Key, null);
							}
						}
					}
					entity.Save();
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
		}

		/// <summary>
		/// Добавить сущность
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="fieldValues"></param>
		/// <returns></returns>
		public static Guid InsertEntity(UserConnection userConnection, string entityName,
			Dictionary<string, object> fieldValues)
		{
			Guid result = Guid.Empty;
			EntitySchema entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			Entity entity = entitySchema.CreateEntity(userConnection);
			entity.SetDefColumnValues();
			try
			{
				foreach (var fieldKeyValue in fieldValues)
				{
					object value = fieldKeyValue.Value;
					string key = fieldKeyValue.Key;
					if (value is string)
					{
						string strValue = value.ToString();
						if (!string.IsNullOrEmpty(strValue))
						{
							var textDataValueType =
								(TextDataValueType) entity.Schema.Columns.GetByColumnValueName(key).DataValueType;
							int size = textDataValueType.Size;
							if (!textDataValueType.IsMaxSize && (strValue.Length > size))
							{
								strValue = strValue.Substring(0, size);
							}
							entity.SetColumnValue(key, strValue);
						}
					}
					else if (value is DateTime)
					{
						if ((DateTime) value != DateTime.MinValue)
						{
							entity.SetColumnValue(key, value);
						}
					}
					else if (value is byte[])
					{
						if ((byte[]) value != new byte[0])
						{
							entity.SetBytesValue(key, (byte[]) value);
						}
					}
					else if (value is Guid)
					{
						if ((Guid) value != Guid.Empty)
						{
							entity.SetColumnValue(key, value);
						}
					}
					else
					{
						if (value != null)
						{
							entity.SetColumnValue(key, value);
						}
					}
				}
				entity.Save();
				result = entity.PrimaryColumnValue;
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return result;
		}
		
		/// <summary>
		/// Получить сущность
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="id"></param>
		/// <returns></returns>
		public static Entity ReadEntity(UserConnection userConnection, string entityName, Guid id)
		{
			var entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			var entity = entitySchema.CreateEntity(userConnection);
			if (entity.FetchFromDB(id))
			{
				return entity;
			}
			return null;
		}

		/// <summary>
		/// Получить коллекцию сущностей по условию
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="conditions"></param>
		/// <returns></returns>
		public static EntityCollection ReadEntityCollection(UserConnection userConnection, string entityName, string[] columnsToFetch, Dictionary<string, object> conditions)
		{
			var esq = new EntitySchemaQuery(userConnection.EntitySchemaManager, entityName);
			esq.PrimaryQueryColumn.IsAlwaysSelect = true;
			foreach (var column in columnsToFetch)
			{
				esq.AddColumn(column);
			}
			foreach (var filter in conditions)
			{
				if (filter.Value == null) esq.Filters.Add(esq.CreateIsNullFilter(filter.Key));
				else esq.Filters.Add(esq.CreateFilterWithParameters(FilterComparisonType.Equal, filter.Key, filter.Value));
			}
			return esq.GetEntityCollection(userConnection);
		}

		/// <summary>
		/// Получить список Id с помощью Select
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="conditions"></param>
		/// <returns></returns>
		public static List<T> ReadRecordsList<T>(UserConnection userConnection, string entityName, string conditionColumn, Dictionary<string, object> conditions)
		{
			List<T> list = new List<T>();
			try
			{
				EntitySchema entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
				Entity entity = entitySchema.CreateEntity(userConnection);
				var select = new Select(userConnection).From(entityName) as Select;
				select.Column(conditionColumn);
				foreach (KeyValuePair<string, object> key in conditions)
				{
					var indexOf = conditions.Keys.ToList().IndexOf(key.Key);
					if (indexOf == 0)
						select.Where(key.Key).IsEqual(Column.Parameter(key.Value));
					else
						select.And(key.Key).IsEqual(Column.Parameter(key.Value));
				}
				using (DBExecutor dbExecutor = userConnection.EnsureDBConnection())
				{
					using (IDataReader reader = select.ExecuteReader(dbExecutor))
					{
						while (reader.Read())
						{
							list.Add(reader.GetColumnValue<T>(conditionColumn));
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return list;
		}

		/// <summary>
		/// Найти сущность
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="id"></param>
		/// <returns></returns>
		public static Entity FindEntity(UserConnection userConnection, string entityName, string conditionName, object conditionValue, string[] columnsToFetch = null)
		{
			var entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			var entity = entitySchema.CreateEntity(userConnection);
			if (columnsToFetch != null)
			{
				if (entity.FetchFromDB(conditionName, conditionValue, columnsToFetch))
				{
					return entity;
				}
			}
			else
			{
				if (entity.FetchFromDB(conditionName, conditionValue))
				{
					return entity;
				}
			}
			return null;
		}

		/// <summary>
		/// Найти Id сущности
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="id"></param>
		/// <returns></returns>
		public static Guid FindEntityId(UserConnection userConnection, string entityName, string conditionName, object conditionValue, string[] columnsToFetch)
		{
			Guid recordId = Guid.Empty;
			var entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			var entity = entitySchema.CreateEntity(userConnection);
			if (entity.FetchFromDB(conditionName, conditionValue, columnsToFetch))
			{
				recordId = entity.GetTypedColumnValue<Guid>("Id");
			}
			return recordId;
		}

		/// <summary>
		/// Добавить запись
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="fieldValues"></param>
		/// <returns></returns>
		public static Guid InsertRecord(UserConnection userConnection, string entityName,
			Dictionary<string, object> fieldValues)
		{
			Guid result = Guid.NewGuid();
			EntitySchema entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
			Entity entity = entitySchema.CreateEntity(userConnection);
			try
			{
				var insert = new Insert(userConnection)
					.Into(entityName);
				if (!fieldValues.ContainsKey("Id"))
					insert.Set("Id", Column.Parameter(result));
				foreach (var fieldValue in fieldValues)
				{
					object value = fieldValue.Value;
					string key = fieldValue.Key;
					if (value is string)
					{
						string strValue = value.ToString();
						if (!string.IsNullOrEmpty(strValue))
						{
							var textDataValueType =
								(TextDataValueType)entity.Schema.Columns.GetByColumnValueName(key).DataValueType;
							int size = textDataValueType.Size;
							if (!textDataValueType.IsMaxSize && (strValue.Length > size))
							{
								strValue = strValue.Substring(0, size);
							}
							insert.Set(fieldValue.Key, Column.Parameter(strValue));
						}
					}
					else if (value is DateTime)
					{
						if ((DateTime)value != DateTime.MinValue)
						{
							insert.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else if (value is Guid)
					{
						if ((Guid)value != Guid.Empty)
						{
							insert.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else
					{
						if (value != null)
						{
							insert.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
				}
				insert.Execute();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return result;
		}

		/// <summary>
		/// Обновить запись
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="primaryColumnValue"></param>
		/// <param name="fieldValues"></param>
		/// <returns></returns>
		public static int UpdateRecord(UserConnection userConnection, string entityName, Guid primaryColumnValue,
			Dictionary<string, object> fieldValues)
		{
			try
			{
				EntitySchema entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
				Entity entity = entitySchema.CreateEntity(userConnection);
				var update = new Update(userConnection, entityName) as Update;
				foreach (var fieldValue in fieldValues)
				{
					object value = fieldValue.Value;
					string key = fieldValue.Key;
					if (value is string)
					{
						string strValue = value.ToString();
						if (!string.IsNullOrEmpty(strValue))
						{
							var textDataValueType =
								(TextDataValueType) entity.Schema.Columns.GetByColumnValueName(key).DataValueType;
							int size = textDataValueType.Size;
							if (!textDataValueType.IsMaxSize && (strValue.Length > size))
							{
								strValue = strValue.Substring(0, size);
							}
							update.Set(fieldValue.Key, Column.Parameter(strValue));
						}
					}
					else if (value is DateTime)
					{
						if ((DateTime) value != DateTime.MinValue)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else if (value is Guid)
					{
						if ((Guid) value != Guid.Empty)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else
					{
						if (value != null)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
						else update.Set(fieldValue.Key, Column.Const(null));
					}
				}
				if (primaryColumnValue != Guid.Empty) update.Where("Id").IsEqual(Column.Parameter(primaryColumnValue));
				return update.Execute();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// Обновить запись
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="conditions"></param>
		/// <param name="fieldValues"></param>
		/// <returns></returns>
		public static int UpdateRecord(UserConnection userConnection, string entityName, Dictionary<string, object> conditions,
			Dictionary<string, object> fieldValues)
		{
			try
			{
				EntitySchema entitySchema = userConnection.EntitySchemaManager.GetInstanceByName(entityName);
				Entity entity = entitySchema.CreateEntity(userConnection);
				var update = new Update(userConnection, entityName) as Update;
				foreach (var fieldValue in fieldValues)
				{
					object value = fieldValue.Value;
					string key = fieldValue.Key;
					if (value is string)
					{
						string strValue = value.ToString();
						if (!string.IsNullOrEmpty(strValue))
						{
							var textDataValueType =
								(TextDataValueType)entity.Schema.Columns.GetByColumnValueName(key).DataValueType;
							int size = textDataValueType.Size;
							if (!textDataValueType.IsMaxSize && (strValue.Length > size))
							{
								strValue = strValue.Substring(0, size);
							}
							update.Set(fieldValue.Key, Column.Parameter(strValue));
						}
					}
					else if (value is DateTime)
					{
						if ((DateTime)value != DateTime.MinValue)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else if (value is Guid)
					{
						if ((Guid)value != Guid.Empty)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
					else
					{
						if (value != null)
						{
							update.Set(fieldValue.Key, Column.Parameter(fieldValue.Value));
						}
					}
				}
				foreach (KeyValuePair<string, object> key in conditions)
				{
					var indexOf = conditions.Keys.ToList().IndexOf(key.Key);
					if (indexOf == 0)
						update.Where(key.Key).IsEqual(Column.Parameter(key.Value));
					else
						update.And(key.Key).IsEqual(Column.Parameter(key.Value));
				}
				return update.Execute();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// Удалить запись
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName"></param>
		/// <param name="credentials"></param>
		/// <returns></returns>
		public static int DeleteRecord(UserConnection userConnection, string entityName, Dictionary<string, object> credentials)
		{
			try
			{
				var delete = new Delete(userConnection)
					.From(entityName) as Delete;
				foreach (KeyValuePair<string, object> key in credentials)
				{
					var indexOf = credentials.Keys.ToList().IndexOf(key.Key);
					if (indexOf == 0)
						delete.Where(key.Key).IsEqual(Column.Parameter(key.Value));
					else
						delete.And(key.Key).IsEqual(Column.Parameter(key.Value));
				}
				return delete.Execute();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// Вставить запись если ее не было
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName">Название схемы</param>
		/// <param name="conditionValues">Поля условия поиска</param>
		/// <param name="fieldValues">Поля заполнения</param>
		public static Guid InsertIfNotRecord(UserConnection userConnection, string entityName, Dictionary<string, object> conditionValues,
			Dictionary<string, object> fieldValues)
		{
			var resultId = SelectEntityId(userConnection, entityName, conditionValues);
			if (resultId == Guid.Empty)
				resultId = InsertRecord(userConnection, entityName, fieldValues);
			return resultId;
		}

		/// <summary>
		/// Вставить запись если ее не было
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName">Название схемы</param>
		/// <param name="conditionValues">Поля условия поиска</param>
		/// <param name="fieldValues">Поля заполнения</param>
		public static Guid InsertIfNotEntity(UserConnection userConnection, string entityName, Dictionary<string, object> conditionValues,
			Dictionary<string, object> fieldValues)
		{
			var resultId = SelectEntityId(userConnection, entityName, conditionValues);
			if (resultId == Guid.Empty)
				resultId = InsertEntity(userConnection, entityName, fieldValues);
			return resultId;
		}

		/// <summary>
		/// Вставить запись если ее не было, иначе - обновить
		/// </summary>
		/// <param name="userConnection"></param>
		/// <param name="entityName">Название схемы</param>
		/// <param name="conditionValues">Поля условия поиска</param>
		/// <param name="fieldValues">Поля заполнения</param>
		public static Guid InsertOrUpdateEntity(UserConnection userConnection, string entityName, Dictionary<string, object> conditionValues,
			Dictionary<string, object> fieldValues)
		{
			var resultId = SelectEntityId(userConnection, entityName, conditionValues);
			if (resultId == Guid.Empty)
				resultId = InsertEntity(userConnection, entityName, fieldValues);
			else UpdateEntity(userConnection, entityName, resultId, fieldValues);
			return resultId;
		}

		public static bool IsValueChanged(object newValue, object oldValue)
		{
			if (newValue == null)
				return false;
			if (newValue is string)
			{
				if (!string.IsNullOrEmpty(newValue.ToString()) && (oldValue == null || newValue.ToString() != oldValue.ToString()))
				{
					return true;
				}
			}
			else if (newValue is DateTime)
			{
				if ((DateTime)newValue != DateTime.MinValue && newValue != oldValue)
				{
					return true;
				}
			}
			else if (newValue is Guid)
			{
				if ((Guid)newValue != Guid.Empty && newValue != oldValue)
				{
					return true;
				}
			}
			else if (newValue is Int32)
			{
				if ((Int32)newValue != 0 && newValue != oldValue)
				{
					return true;
				}
			}
			return false;
		}
	}
}