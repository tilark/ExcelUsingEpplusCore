using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelUsingEpplusCore.ViewModels.ReadFromExcel;
using Microsoft.AspNetCore.Http;
using System.Text;
using System.IO;
using System.Reflection;
using ExcelWithEpplusCore452;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Net.Http.Headers;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExcelUsingEpplusCore.Controllers
{
    public class TestReadFromExcelController : Controller
    {
        private IHostingEnvironment hostingEnv;
        private string uploadFileDirectory = String.Empty;
        public TestReadFromExcelController(IHostingEnvironment hostingEnv)
        {
            this.hostingEnv = hostingEnv;
            uploadFileDirectory = "/UploadFiles";
        }
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
                ViewBag.ErrorMessage = "上传文件不符合规范，只接受xlsx类型！";

                return View();
            }
            Dictionary<string, DepartmentFromExcelViewModel> viewModel = new Dictionary<string, DepartmentFromExcelViewModel>();
            
            var fileStream = GetFileStream(files);
            viewModel = GetDepartmentMessageFromExcel(fileStream);

            if (viewModel == null)
            {
                ViewBag.ErrorMessage = "读取文件内容失败！";
            }
            //将读取的内容显示到页面
            return View(viewModel);
        }
        private Dictionary<string, DepartmentFromExcelViewModel> GetDepartmentMessageFromExcel(Stream fileStream)
        {
            var factory = ExcelEntityFactory.GetInstance();
            var TPropertyNameDisplayAttributeNameDic = GetTPropertyNameDisplayAttributeNameDic<DepartmentFromExcelViewModel>();
            var result = factory.CreateReadFromExcel().ExcelToEntityDictionary<DepartmentFromExcelViewModel>(TPropertyNameDisplayAttributeNameDic, fileStream, out StringBuilder errorMesg);
            ViewBag.ErrorMessage = errorMesg.ToString();
            return result;
        }


        /// <summary>
        /// 文件存在，且文件名是以xlsx结尾
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        private bool IsValidFile(List<IFormFile> files)
        {
            bool result = false;
            if (files != null && files.Count > 0)
            {
               if(files.All(a => a.ContentType.Equals(ExcelEntityFactory.GetInstance().Excel2007ContentType)))
                {
                    result = true;
                }
            }
            return result;
        }

        /// <summary>
        /// 获得第一个上传文件的Stream
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
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

        #region 上传到服务器目录
        private string GetSavedFilePathInServer(List<IFormFile> files)
        {
            var file = files.First();
            var fileName = ContentDispositionHeaderValue
                               .Parse(file.ContentDisposition)
                               .FileName
                               .Trim('"');
            string fileDirectory = GetFileDirectory();
            bool savedResult = CreateOrRenameIfExsis(fileDirectory, fileName, file, out string finalFileName);
            return finalFileName;
        }


        private Dictionary<string, DepartmentFromExcelViewModel> GetDepartmentMessageFromExcelWithFilePath(string filePath)
        {
            var factory = ExcelEntityFactory.GetInstance();
            var TPropertyNameDisplayAttributeNameDic = GetTPropertyNameDisplayAttributeNameDic<DepartmentFromExcelViewModel>();
            var result = factory.CreateReadFromExcel().ExcelToEntityDictionary<DepartmentFromExcelViewModel>(TPropertyNameDisplayAttributeNameDic, filePath, out StringBuilder errorMesg);
            ViewBag.ErrorMessage = errorMesg.ToString();
            return result;
        }
        /// <summary>
        /// 创建或重命名文件，返回最终的文件名
        /// </summary>
        /// <param name="fileDirectory">文件上传的目录</param>
        /// <param name="fileName">文件名</param>
        /// <param name="file">文件信息,file.filename不可靠</param>
        /// <returns>返回在服务器上的文件名</returns>
        private bool CreateOrRenameIfExsis(string fileDirectory, string fileName, IFormFile file, out string finalFilePath)
        {

            bool result = false;
            try
            {
                //若目录不存在，创建新的目录
                if (Directory.Exists(fileDirectory) == false)
                {
                    Directory.CreateDirectory(fileDirectory);
                }
                string tempFileName = fileName.Split('.')[0];
                string tempFileType = fileName.Split('.')[1];
                fileDirectory = fileDirectory + "\\";
                var filePath = fileDirectory + fileName;
                //如果文件存在，重命名文件名
                int i = 1;
                while (System.IO.File.Exists(filePath))
                {
                    fileName = tempFileName + "(" + i.ToString() + ")" + "." + tempFileType;
                    filePath = fileDirectory + fileName;
                    i++;
                    //System.IO.File.Delete(fileUrl);
                }
                using (FileStream fs = System.IO.File.Create(filePath))
                {
                    file.CopyTo(fs);
                    fs.Flush();
                }
                finalFilePath = filePath;

                result = true;
            }
            catch (Exception)
            {

                throw;
            }
            return result;
        }

        private string GetFileDirectory()
        {
            string fileDirectory = hostingEnv.WebRootPath + uploadFileDirectory.Replace("/", "\\");
            return fileDirectory;
        }

        #endregion
    }
}
