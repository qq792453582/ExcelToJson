using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelToJson
{
	public class ExcelConverter
	{
		private static readonly Dictionary<string, object> s_GlobalTypeRegedit = new Dictionary<string, object>();

		private static readonly HashSet<string> s_TrueValues = new HashSet<string>
		{
			"是",
			"1",
			"True",
			"true"
		};

		private readonly Dictionary<string, object> m_LocalTypeRegedit = new Dictionary<string, object>();

		static ExcelConverter()
		{
			RegisterGlobalType("string", ValueToString);
			RegisterGlobalType("number", ValueToNumber);
			RegisterGlobalType("bool", ValueToBool);
			RegisterGlobalType("int", ValueToInteger);
		}

		public JObject data { get; } = new JObject();

		public static object ValueToString(string value)
		{
			return value;
		}

		public static object ValueToBool(string value)
		{
			return s_TrueValues.Contains(value);
		}

		public static object ValueToNumber(string value)
		{
			double.TryParse(value, out var number);
			return number;
		}

		public static object ValueToInteger(string value)
		{
			int.TryParse(value, out var number);
			return number;
		}

		public static void RegisterGlobalType(string typeName, Func<string, object?> typeConverter)
		{
			s_GlobalTypeRegedit.Add(typeName, typeConverter);
		}

		public static void RegisterGlobalType(string typeName, Dictionary<string, object> typeMap)
		{
			s_GlobalTypeRegedit.Add(typeName, typeMap);
		}

		public ExcelReader ReadExcel(string filePath)
		{
			Console.WriteLine(Path.GetFullPath(filePath));
			using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				using (var excelDataReader = ExcelReaderFactory.CreateReader(stream))
				{
					var dataSet = excelDataReader.AsDataSet(new ExcelDataSetConfiguration
					{
						UseColumnDataType = false,
						ConfigureDataTable = _ => new ExcelDataTableConfiguration
						{
							UseHeaderRow = true
						}
					});
					return new ExcelReader(this, dataSet);
				}
			}
		}

		public void RegisterLocalType(string typeName, Func<string?, object> typeConverter)
		{
			m_LocalTypeRegedit.Add(typeName, typeConverter);
		}

		public void RegisterLocalType(string typeName, Dictionary<string, object> typeMap)
		{
			m_LocalTypeRegedit.Add(typeName, typeMap);
		}

		public object? ParseValue(string? typeName, string? value)
		{
			if (string.IsNullOrEmpty(typeName)) return null;

			if (!m_LocalTypeRegedit.TryGetValue(typeName, out var typeConverter)) s_GlobalTypeRegedit.TryGetValue(typeName, out typeConverter);

			if (typeConverter is Func<string?, object> typeConvertFunc) return typeConvertFunc(value);

			if (!string.IsNullOrEmpty(value))
			{
				if (typeConverter is Dictionary<string, object> typeMap) return typeMap.GetValueOrDefault(value);
			}

			return null;
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
