# TitleAttribute

不依赖任何字段，指的是：

```CSharp
// Title 标记的位置基本不影响实际效果，只为了方便阅读
internal class Test1
{
    [Title("姓名：", "B2")]
    public string StudentName { get; set; }

    [Title("性别：", "D2")]
    public string Sex { get; set; }
}

// 效果一样，并不会因为标记在不同的字段上就会有所差别
internal class Test2
{
    public string StudentName { get; set; }

    [Title("姓名：", "B2")]
    [Title("性别：", "D2")]
    public string Sex { get; set; }
}
```