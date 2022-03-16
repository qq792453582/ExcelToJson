using Newtonsoft.Json;
using NUnit.Framework;

namespace ExcelToJson.Test;

[TestFixture]
public class ExcelConvertTest
{
	[Test]
	public void TestReadExcel()
	{
		var convert = new ExcelConverter();
		using (var excelReader = convert.ReadExcel("Test/Data/建筑.xlsx"))
		{
			excelReader.ReadSheet("标签页")?.Convert().Apply("builds");
		}

		Console.WriteLine(JsonConvert.SerializeObject(convert.data));
	}

	[Test]
	public void CodeTest()
	{
	}
}
