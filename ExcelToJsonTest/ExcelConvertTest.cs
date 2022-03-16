using System.Text;
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
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

			var converter = new ExcelConverter();

			using (var excelReader = converter.ReadExcel("Data/建筑.xlsx"))
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

				excelReader.ReadSheet("显示规则")?.Convert(data =>
				{
					var typeMap = data.ToObject<Dictionary<string, object>>();
					if (typeMap != null) converter.RegisterLocalType("buildDisplay", typeMap);
					return data;
				});

				excelReader.ReadSheet("建筑")?.Convert().Apply("builds");
			}

			Console.WriteLine(JsonConvert.SerializeObject(converter.data, Formatting.Indented));
			Assert.IsTrue(JToken.DeepEquals(converter.data, JsonConvert.DeserializeObject<JObject>(
				@"{
					buildPages: [
						{
							name: '基础'
						}
					],
					buildGroups: {
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
					},
					builds: {
						'1001': {
							id: 1001,
							group: 1,
							name: '墙`道源',
							order: 1,
							icon: 'slg_build_well_2',
							effect: '',
							displayRule: 'well',
							editorOnly: false,
						}
					}
				}")));
		}
	}
}
