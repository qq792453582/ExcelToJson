using System;
using System.Data;

namespace ExcelToJson
{
	public class ExcelReader : IDisposable
	{
		private readonly ExcelConverter m_Converter;
		private readonly DataSet m_DataSet;

		public ExcelReader(ExcelConverter converter, DataSet dataSet)
		{
			m_Converter = converter;
			m_DataSet = dataSet;
		}

		public void Dispose()
		{
			m_DataSet.Dispose();
		}

		public SheetConverter? ReadSheet(string sheetName)

		{
			var dataTable = m_DataSet.Tables[sheetName];

			if (dataTable != null)
			{
				return new SheetConverter(m_Converter, dataTable);
			}


			return null;
		}
	}
}
