using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using JsonSerializer = System.Text.Json.JsonSerializer;

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

				excelReader.ReadSheet("建筑组")?.Convert().Apply("buildGroups");
			}

			Console.WriteLine(JsonConvert.SerializeObject(converter.data));
			Assert.IsTrue(JToken.DeepEquals(converter.data, JsonConvert.DeserializeObject<JObject>(
				@"{
					buildPages:[{
							name: '基础'
					}],
					buildGroups:{
						'1': {
							id: 1,
							name: '墙',
							order: 1,
							page: 0,
							mergeInGroup: 0,
							usePaint: true,
							useErase: true,
							editorOnly: false,
						}
					}
				}")));
		}

		[Test]
		public void CodeTest()
		{
			Console.WriteLine(JsonConvert.SerializeObject(JsonConvert.DeserializeObject(
				@"{
					buildPages:[{
							name: '基础'
					}],
					buildGroups:{
						1: {
							id: 1,
							name: '墙',
							order: 1,
							page: 0,
							mergeInGroup: 0,
						}
					}
				}")));
		}
	}
}
