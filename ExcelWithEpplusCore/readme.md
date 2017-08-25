# ExcelWithEpplusCore
## 概述
使用EpplusCore插件导入/导出Excel
## 导出Excel
1. Controller/Action
```
 var factory = ExcelEntityFactory.GetInstance();
           //测试数据列表
            var testData = GetTestData();
        //获得标题名，获得属性中DisplayAttribute，如果不存在，则使用属性名
            var titleHead = GetTitleList();
            var fileContent = factory.CreateWriteToExcel().ExportListToExcel<TestA>(testData, titleHead, true);
 return File(fileContent, factory.CreateWriteToExcel().ExcelContentType, System.DateTime.Now.ToString("yyyyMMddHHmmssffff")+".xlsx");
```
2. GetTestData()
以TestA为类型：
```
 public class TestA
        {
            [Required]
            [Display(Name = "资产来源分类组名称")]

            public virtual string Name { get; set; }

            [Display(Name = "排序号码")]
            public virtual long SortNumber { get; set; }

            [Display(Name = "备注")]
            public virtual string Remarks { get; set; }
        }
 private List<TestA> GetTestData()
        {
            var result = new List<TestA>();
            for(int i = 0; i<5; i++)
            {
                var temp = new TestA
                {
                    Name = "Test" + i.ToString(),
                    SortNumber = i,
                    Remarks = "Remarks" + i.ToString()
                };
                result.Add(temp);
            }
            return result;
        }
```
3. GetTitleList()
```
 private List<string> GetTitleList<T>() where T: class
        {
            var result = new List<string>();
            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                var displayName = property.GetCustomAttributes(typeof(DisplayAttribute), true).FirstOrDefault() as DisplayAttribute;
                if (displayName == null)
                {
                    result.Add(property.Name);
                }
                else
                    result.Add((displayName).GetName());
            }
            return result;
        }
```
