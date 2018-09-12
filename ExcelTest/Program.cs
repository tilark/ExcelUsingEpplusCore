using System;
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
    }
}
