## 渲染示例

渲染表单
```CSharp
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