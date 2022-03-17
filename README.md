# ExcelToJson
 Excel Convert To Json Object

这是一个将Excel表格转换为Json Object的转换器，你可以拿到你想要的任何结构的数据对象

此转换器并附带转换为Json文本的方法， 你也可以很方便的拓展这个转换器，以你想要的方式导出为任意格式

## 如何使用
```c#
var converter = new ExcelConverter();
using (var excelReader = converter.ReadExcel(filePath))
{
	excelReader.ReadSheet(sheetName)?.Convert(postprocessor).Apply(key);
}
```

## 自定义类型
```c#
converter.RegisterLocalType(typeName, tpyeConveter);
```

## 表格配置

|   ID   |    名字    |     标记1     |     标记2     |
|:------:|:--------:|:-----------:|:-----------:|
|        | &ID.name | &ID.mark.#0 | &ID.mark.#1 |
| number |  string  |   number    |   number    |
|   1    |    熊大    |      1      |      2      |
|   2    |    熊二    |      3      |      4      |

### 输出
```json
{
	"1": {
		"name": "熊大",
		"mark": [1, 2]
	},
	"2": {
		"name": "熊二",
		"mark": [3, 4]
	}
}
```

## 字段标识

### \# 数组索引标记，父对象将自动生成为数组
### & 引用标记，该字段值将自动转换为同一行对应列的单元格的值
### . 分割标记，可将结构体进行嵌套

## [示例](ExcelToJsonTest)
