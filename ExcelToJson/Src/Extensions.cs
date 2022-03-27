using System.Data;
using Newtonsoft.Json.Linq;

namespace ExcelToJson
{
	internal static class Extensions
	{
		public static object? ElementAt(this DataTable dataTable, int rowsNum, int columnsNumber)
		{
			if (dataTable.Rows.Count > rowsNum)
			{
				return dataTable.Rows[rowsNum][columnsNumber];
			}

			return null;
		}

		public static object? ElementAt(this DataTable dataTable, int rowsNum, string columnsKey)
		{
			if (dataTable.Rows.Count > rowsNum)
			{
				return dataTable.Rows[rowsNum][columnsKey];
			}

			return null;
		}

		public static bool IsNullOrEmpty(JToken? token)
		{
			return token == null || token.Type == JTokenType.Null;
		}
	}
}
