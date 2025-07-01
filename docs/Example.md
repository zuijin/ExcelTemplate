# 使用示例

## 实体类示例

```Csharp
public class MyModel
{
    [Title("姓名：", "B2")]
    [Position("C2")]
    public string StudentName { get; set; }

    [Title("性别：", "D2")]
    [Position("E2")]
    public string Sex { get; set; }

    [Title("生日：", "B3")]
    [Position("C3")]
    public DateTime BirthDate { get; set; }
}
```

## 读取模板示例

```Csharp
var template = Template.FromType<MyModel>();
```

详见 [模板生成](TemplateExample.md)



## 读取 Excel 示例

```Csharp
// excel 数据文件
// 先转换为流形式
var dataExcelFileName = "...";
using var stream = File.OpenRead(dataExcelFileName);

// 读取数据
T data = template.Capture<T>(stream);
```

## 渲染 Excel 示例

渲染并返回 IWorkbook 对象

```CSharp
IWorkbook workbook = template.Render(data);
```

向已有 Excel 文件渲染数据

```CSharp
var excelFileName = "...";
var excelFileStream = File.OpenWrite(excelFileName);
var workbook = WorkbookFactory.Create(excelFileStream);

// 将数据渲染到 workbook 对象
template.Render(data, workbook);

// 或者，直接渲染到流对象
template.Render(data, excelFileStream);
```