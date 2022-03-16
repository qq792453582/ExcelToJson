using System.Data;

namespace ExcelToJson
{
	public static class DataTableExtensions
	{
		public static object? ElementAt(this DataTable dataTable, int rowsNum, int columnsNumber)
		{
			if (dataTable.Rows.Count > rowsNum) return dataTable.Rows[rowsNum][columnsNumber];

			return null;
		}

		public static object? ElementAt(this DataTable dataTable, int rowsNum, string columnsKey)
		{
			if (dataTable.Rows.Count > rowsNum) return dataTable.Rows[rowsNum][columnsKey];

			return null;
		}
	}
}
