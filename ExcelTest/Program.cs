using System;
using System.Collections.Generic;
using System.IO;
namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            TestExcelToDicitonary();
        }
        static void TestExcelToDicitonary()
        {
            using (var file = File.Open("e:/test.xlsx", FileMode.Open))
            {
                var readExcel = ExcelWithEpplusCore.ExcelEntityFactory.GetInstance().CreateReadFromExcel();
                var readResult = readExcel.ExcelToDicitonary(file);
                int key = 1;

                foreach (var firstDic in readResult)
                {
                    Console.WriteLine("Key:" + key);
                    Console.WriteLine(firstDic.Key);
                    foreach(var secondDic in firstDic.Value)
                    {
                        Console.WriteLine("Key: " + secondDic.Key + ", " + "Value: " + secondDic.Value);

                    }
                    key++;
                }
            }
        }

        static void TestExportDictionaryToExcel()
        {
            var writeExcel = ExcelWithEpplusCore.ExcelEntityFactory.GetInstance().CreateWriteToExcel();
            var testData = DictionaryTestData();
            var result = writeExcel.ExportDictionaryToExcel(testData, false);
            File.WriteAllBytes("e:/testwrite2", result);  
        }

        static private Dictionary<string, Dictionary<string, string>> DictionaryTestData()
        {
            var result = new Dictionary<string, Dictionary<string, string>>();
            var secondDic1 = new Dictionary<string, string>();
            secondDic1.Add("姓名", "张三");
            secondDic1.Add("部门", "人事部");
            secondDic1.Add("基本工资", "15000");
            result.Add("201808", secondDic1);
            result.Add("201809", secondDic1);
            return result;
        }
    }
}
