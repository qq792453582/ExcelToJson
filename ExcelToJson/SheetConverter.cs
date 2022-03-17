using System;
using System.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelToJson
{
	public class SheetConverter
	{
		private readonly ExcelConverter m_Converter;

		private readonly DataTable m_DataTable;

		public SheetConverter(ExcelConverter converter, DataTable dataTable)
		{
			m_DataTable = dataTable;
			m_Converter = converter;
		}

		public JToken? data { get; private set; }

		public SheetConverter Convert(Func<JToken, JToken>? postprocessor = null)
		{
			for (var columnsNumber = 0; columnsNumber < m_DataTable.Columns.Count; columnsNumber++)
			{
				var propertyKey = m_DataTable.ElementAt(0, columnsNumber)?.ToString();
				var typeName = m_DataTable.ElementAt(1, columnsNumber)?.ToString();

				for (var rowsNum = 2; rowsNum < m_DataTable.Rows.Count; rowsNum++)
				{
					var propertyKeyTokens = GetPropertyKeys(propertyKey, rowsNum);
					if (propertyKeyTokens != null)
					{
						var value = m_DataTable.ElementAt(rowsNum, columnsNumber)?.ToString();
						var result = m_Converter.ParseValue(typeName, value);

						if (result != null) SetValue(propertyKeyTokens, result);
					}
				}
			}

			if (postprocessor != null && data != null) data = postprocessor(data);

			return this;
		}

		private void SetValue(string[] propertyKeys, object value)
		{
			data = SetValue(data, propertyKeys, 0, value);
		}

		private JToken? SetValue(JToken? target, string[] propertyKeys, int depth, object value)
		{
			var propertyKey = propertyKeys[depth];

			var isIndex = propertyKey.StartsWith("#");
			if (isIndex)
			{
				if (int.TryParse(propertyKey.Substring(1), out var index))
				{
					if (Extensions.IsNullOrEmpty(target)) target = new JArray();
					if (target is JArray array)
					{
						for (var i = array.Count; i <= index; i++) array.Add(null!);
						;
						if (depth == propertyKeys.Length - 1)
						{
							array[index] = new JValue(value);
						}
						else
						{
							var token = SetValue(array[index], propertyKeys, depth + 1, value);
							if (token != null) array[index] = token;
						}
					}
				}
			}
			else
			{
				if (Extensions.IsNullOrEmpty(target)) target = new JObject();

				if (target is JObject obj)
				{
					if (depth == propertyKeys.Length - 1)
					{
						obj[propertyKey] = new JValue(value);
					}
					else
					{
						var token = SetValue(obj[propertyKey], propertyKeys, depth + 1, value);
						if (token != null) obj[propertyKey] = token;
					}
				}
			}

			return target;
		}

		public void Apply(string key)
		{
			if (data != null) m_Converter.data.Add(key, data);
		}

		public void ApplyToType(string typeName)
		{
			if (data is JObject obj)
			{
				m_Converter.RegisterLocalType(typeName, value =>
				{
					if (string.IsNullOrEmpty(value))
					{
						return null;
					}

					return obj[value]?.ToObject<object>();
				});
			}
			else if (data is JArray array)
			{
				m_Converter.RegisterLocalType(typeName, value =>
				{
					if (int.TryParse(value, out var index))
					{
						if (array.Count > index)
						{
							return array[index].ToObject<object>();
						}
					}

					return null;
				});
			}
		}

		private string[]? GetPropertyKeys(string? propertyPath, int rowsNumber)
		{
			if (string.IsNullOrEmpty(propertyPath)) return null;

			var keys = propertyPath.Split(".");
			for (var i = 0; i < keys.Length; i++)
			{
				var key = keys[i];
				var isIndex = key.StartsWith("#");
				if (isIndex) key = key.Substring(1);

				if (key.StartsWith("&"))
				{
					key = key.Substring(1);
					var value = m_DataTable.ElementAt(rowsNumber, key)?.ToString();
					if (string.IsNullOrEmpty(value)) return null;

					var typeName = m_DataTable.ElementAt(1, key)?.ToString();
					key = m_Converter.ParseValue(typeName, value)?.ToString();
				}

				if (string.IsNullOrEmpty(key)) return null;

				if (isIndex)
				{
					if (int.TryParse(key, out _))
						key = "#" + key;
					else
						return null;
				}

				keys[i] = key;
			}

			return keys;
		}

		public string ToJson()
		{
			return JsonConvert.SerializeObject(data);
		}

		public string ToJson(Formatting formatting)
		{
			return JsonConvert.SerializeObject(data, formatting);
		}

		public string ToJson(Formatting formatting, JsonSerializerSettings? settings)
		{
			return JsonConvert.SerializeObject(data, formatting, settings);
		}

		public string ToJson(JsonSerializerSettings? settings)
		{
			return JsonConvert.SerializeObject(data, settings);
		}
	}
}
