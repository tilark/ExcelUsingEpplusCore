using System;
using System.IO;

namespace ExcelUsingEpplusCoreConsoleTest
{
    /// <summary>
    /// 测试ExcelUsingEpplusCore
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            //读取Excel文件

            Console.WriteLine("Hello World!");
            TestExcelToDicitonary();
        }

        static void TestExcelToDicitonary()
        {
            using (var file = File.Open("e:/test.xlsx", FileMode.Open))
            {
                var factory = ExcelWithEpplusCore.ExcelEntityFactory.GetInstance();
                var readExcel = factory.CreateReadFromExcel().ExcelToDicitonary(file);
                int firstKey = 1;
                foreach(var firstDic in readExcel)
                {
                    Console.WriteLine(firstKey + ":");
                    Console.WriteLine(firstDic.Key);
                    foreach(var secondDic in firstDic.Value)
                    {
                        Console.WriteLine("key:" + secondDic.Key + ";" + "value:" + secondDic.Value);
                    }
                    firstKey++;
                }
            }
        }
    }
}
