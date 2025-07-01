# 从类定义生成模板

## 类定义

先定义实体类：

```CSharp

internal class MixtureModel
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

    [Title("第一次模拟考试", "B5", "E5")]
    [Position("B6")]
    public List<ScoreItem> Scores_1st { get; set; }

    [Title("总分：", "D10")]
    [Position("E10")]
    public float TotalScore_1st { get; set; }

    [Title("总排名：", "D11")]
    [Position("E11")]
    public int TotalRanking_1st { get; set; }

    [Title("第二次模拟考试", "B13", "E13")]
    [Position("B14")]
    public List<ScoreItem> Scores_2nd { get; set; }

    [Title("总分：", "D18")]
    [Position("E18")]
    public float TotalScore_2nd { get; set; }

    [Title("总排名：", "D19")]
    [Position("E19")]
    public int TotalRanking_2nd { get; set; }


    internal class ScoreItem
    {
        [Col("科目", 0)]
        public string SubjectName { get; set; }

        [Merge("成绩")]
        [Col("分数", 1)]
        public float Score { get; set; }

        [Merge("成绩")]
        [Col("排名", 2)]
        public int Ranking { get; set; }

        [Col("考试时间", 3)]
        public DateTime ExamTime { get; set; }
    }
}

```

## 特性 Attribute 解释


上面使用了几个 **Attribute** 对模板内容进行了描述：

|  特性  | 描述 |
|---|---|
| **Title** | 在指定位置输出文本，这个特性本身不赖于任何字段，详见 [TitleAttribute](../Attributes/Title.md) |
| **Position** | 给字段设置一个位置值，表示将该字段的值输出到对应的位置 |
| **Col** | 定义一个列，作用于数组类型的字段下，设置 Table 的 **表头文本** 和 **列顺序** |
| **Merge** | 用于合并表头 |


## 读取模板信息

```CSharp
//这样读取模板
var template = Template.FromType(typeof(MixtureModel));

//或者
var template = Template.FromType<MixtureModel>();

//或者
var template = new TypeDesignAnalysis().DesignAnalysis(typeof(MixtureModel));
```

生成模板只是借助了 类字段 和 特性 的描述信息，生成的模板本身并不跟该实体类强绑定，就比如 template 实例，本身并不会跟 MixtureModel 类有任何强绑定关系。所以，实体类只要结构定义符合要求即可。

## 使用样式

模板引擎还支持简单的样式，方便控制 Excel 的渲染格式

```Csharp
[StyleDic(Key = "date_format", DataFormat = "yyyy/m/d h:mm:ss")]
[StyleDic(Key = "title_red", TextColor = "FF0000", FontHeightInPoints = 18, IsBold = true, HorizontalAlignment = ETHorizontalAlignment.Center, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, WrapText = true)]
internal class TestFormStyleModel
{
    [Title("姓名：", "B2", Style = "title_red")]
    [Position("C2", Style = "title_red")]
    public string StudentName { get; set; }

    [Title("性别：", "D2")]
    [Position("E2")]
    [Style(BgColor = "0000FF", FontHeightInPoints = 12, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Left)]
    public string Sex { get; set; }

    [Title("生日：", "B3")]
    [Position("C3", Style = "date_format")]
    public DateTime BirthDate { get; set; }
}
```

|  特性  | 描述 |
|---|---|
| **Style** | 直接标记在字段上，设置该字段的显示样式 |
| **StyleDic** | 标记在类上，通过设置 Key 属性跟其他特性 Style 关联，可以方便多个字段样式复用 |
