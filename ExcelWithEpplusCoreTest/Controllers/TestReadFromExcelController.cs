using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.IO;
using ExcelWithEpplusCoreTest.ViewModels.ReadFromExcel;
using ExcelWithEpplusCore;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace ExcelWithEpplusCoreTest.Controllers
{
    public class TestReadFromExcelController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ImportDepartmentMessageStream()
        {
            return View();
        }
        [HttpPost]
        [AutoValidateAntiforgeryToken]
        public IActionResult ImportDepartmentMessageStream(List<IFormFile> files)
        {
            if (!IsValidFile(files))
            {
                return new EmptyResult();
            }
            Stream fileStream = GetFileStream(files);
            Dictionary<string, DepartmentFromExcelViewModel> viewModel = GetDepartmentMessageFromExcel(fileStream);
            //将读取的内容显示到页面
            return View(viewModel);
        }

        private bool IsValidFile(List<IFormFile> files)
        {
            if (files != null && files.Count > 0)
            {
                return true;
            }
            else { return false; }
        }

        private Dictionary<string, DepartmentFromExcelViewModel> GetDepartmentMessageFromExcel(Stream fileStream)
        {
            var factory = ExcelEntityFactory.GetInstance();
            var TPropertyNameDisplayAttributeNameDic = GetTPropertyNameDisplayAttributeNameDic<DepartmentFromExcelViewModel>();
            var result = factory.CreateReadFromExcel().ExcelToEntityDictionary<DepartmentFromExcelViewModel>(TPropertyNameDisplayAttributeNameDic, fileStream, out StringBuilder errorMesg);
            ViewBag.ErrorMessage = errorMesg.ToString();
            return result;
        }

        private Stream GetFileStream(List<IFormFile> files)
        {
            var file = files.First();
            var fileStream = file.OpenReadStream();
            return fileStream;
        }

        /// <summary>
        /// 获取类中属性的DisplayAttribute值，该值作为Value，属性名作为Key
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected Dictionary<string, string> GetTPropertyNameDisplayAttributeNameDic<T>()
        {
            List<PropertyInfo> propertyInfoList = new List<PropertyInfo>(typeof(T).GetProperties());
            var cellHeader = new Dictionary<string, string>();
            foreach (var property in propertyInfoList)
            {
                //var displayName2 = property.GetCustomAttribute<DisplayNameAttribute>(true).DisplayName;
                cellHeader[property.Name] = (property.GetCustomAttributes(typeof(DisplayAttribute), true).FirstOrDefault() as DisplayAttribute).GetName();
            }

            return cellHeader;
        }
    }
}