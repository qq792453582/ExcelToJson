using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;

namespace ExcelToJson.Test
{
	[TestFixture]
	public class ExcelConvertTest
	{
		[Test]
		public void TestReadExcel()
		{
			var converter = new ExcelConverter();

			using (var excelReader = converter.ReadExcel("Test/Data/建筑.xlsx"))
			{
				var pages = new Dictionary<string, object>();

				excelReader.ReadSheet("标签页")?.Convert(data =>
				{
					if (data is JArray array)
						for (var i = 0; i < array.Count; i++)
						{
							var pageData = array[i];

							if (!Extensions.IsNullOrEmpty(pageData))
							{
								var pageName = pageData.Value<string>("name");
								if (!string.IsNullOrEmpty(pageName)) pages[pageName] = i;
							}
						}

					return data;
				}).Apply("buildPages");

				converter.RegisterLocalType("buildGroupPage", pages);
			}

			Assert.IsTrue(JToken.DeepEquals(converter.data, JsonConvert.DeserializeObject<JObject>(
				@"{
					'buildPages':[{
							'name':'基础'
						}]
				}")));
		}

		[Test]
		public void CodeTest()
		{
		}
	}
}
