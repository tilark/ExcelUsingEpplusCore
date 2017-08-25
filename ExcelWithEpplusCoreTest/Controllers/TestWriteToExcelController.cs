using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelWithEpplusCoreTest.ViewModels;
using ExcelWithEpplusCore;

namespace ExcelWithEpplusCoreTest.Controllers
{
    public class TestWriteToExcelController : Controller
    {
        // GET: TestExcelUsingEpplus
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult TestExportListToExcel()
        {
            var factory = ExcelEntityFactory.GetInstance();

            var testData = GetTestData();
            var titleHead = GetTitleList();
            var result = factory.CreateWriteToExcel().ExportListToExcel<TestA>(testData, titleHead, true);

            return new FileContentResult(result, factory.CreateWriteToExcel().ExcelContentType);

        }
        public ActionResult TestExportListToExcelWithOutHeading()
        {
            var factory = ExcelEntityFactory.GetInstance();

            var testData = GetTestData();
            var titleHead = GetTitleList();
            var result = factory.CreateWriteToExcel().ExportListToExcel<TestA>(testData, null, true);

            return new FileContentResult(result, factory.CreateWriteToExcel().ExcelContentType);
        }
        public ActionResult TestExportListToExcelWithOutisShowSlNoIsFalse()
        {
            var factory = ExcelEntityFactory.GetInstance();

            var testData = GetTestData();
            var titleHead = GetTitleList();
            var result = factory.CreateWriteToExcel().ExportListToExcel<TestA>(testData, null, false);

            return new FileContentResult(result, factory.CreateWriteToExcel().ExcelContentType);
        }
        private List<string> GetTitleList()
        {
            var result = new List<string>();
            result.Add("名称");
            result.Add("排序号码");
            result.Add("备注");
            return result;
        }
        private List<TestA> GetTestData()
        {
            var result = new List<TestA>();
            for (int i = 0; i < 5; i++)
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
    }
}