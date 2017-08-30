# 使用说明
## 在使用时需再增加一个Abstract Class
```
public abstract class ReadFromExcelService
    {
        private string excelType;
        private Stream fileStream;
        private string startCellName;
        private int keyColumn;
        private int mergeTitleRow;


        public ReadFromExcelService(string excelType, Stream fileStream, string startCellName = "B1", int mergeTitleRow = 1, int keyColumn = 1)
        {
            this.excelType = excelType;
            this.fileStream = fileStream;
            this.startCellName = startCellName;
            this.mergeTitleRow = mergeTitleRow;
            this.keyColumn = keyColumn;
        }
        /// <summary>
        /// 验证Excel的格式是否为指定格式
        /// </summary>
        /// <returns></returns>
        protected virtual bool ValidExcelFormat()
        {
            if (excelType.Equals(ExcelEntityFactory.GetInstance().Excel2007ContentType))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 从Excel中读取数据，
        /// </summary>
        /// <typeparam name="T">与Excel对应的数据模型</typeparam>
        /// <param name="fileStream">文件Stream</param>
        /// <param name="startCellName">数据开始的单元格，如Key在A列，标题行只有一行在第1行，则数据开始的单元格为"B1"</param>
        /// <returns></returns>
        protected virtual Dictionary<string, T> GetMessageFromExcelWithFileStream<T>() where T : new()
        {

            var factory = ExcelEntityFactory.GetInstance();
            var TPropertyNameDisplayAttributeNameDic = GetTPropertyNameDisplayAttributeNameDic<T>();
            var result = factory.CreateReadFromExcel().ExcelToEntityDictionary<T>(TPropertyNameDisplayAttributeNameDic, fileStream, out StringBuilder errorMesg, null, startCellName, mergeTitleRow, keyColumn);
            return result;
        }

        /// <summary>
        /// 将Excel表中的数据写入到数据库中，返回出现的错误信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="messageFromExcel"></param>
        /// <returns>返回添加到数据库中出现的错误信息</returns>
        protected abstract Task<string> SaveDataFromExcelToDataBase<T>(Dictionary<string, T> messageFromExcel) where T : new();

        /// <summary>
        /// 执行顺序
        /// </summary>
        public virtual async Task<string> ExcuteAsync<T>() where T : new()
        {
            if (ValidExcelFormat())
            {
                var messageFromExcel = GetMessageFromExcelWithFileStream<T>();
                return await SaveDataFromExcelToDataBase<T>(messageFromExcel);
            }
            else
            {
                return "Excel格式错误，只接受.xlsx的文件格式";
            }
        }
        /// <summary>
        /// 获取类中属性的DisplayAttribute值，该值作为Value，属性名作为Key
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private Dictionary<string, string> GetTPropertyNameDisplayAttributeNameDic<T>()
        {
            List<PropertyInfo> propertyInfoList = new List<PropertyInfo>(typeof(T).GetProperties());
            var propertyNameDisplayAttributeNameDic = new Dictionary<string, string>();
            foreach (var property in propertyInfoList)
            {
                //var displayName2 = property.GetCustomAttribute<DisplayNameAttribute>(true).DisplayName;
                propertyNameDisplayAttributeNameDic[property.Name] = (property.GetCustomAttributes(typeof(DisplayAttribute), true).FirstOrDefault() as DisplayAttribute).GetName();
            }
            return propertyNameDisplayAttributeNameDic;
        }
    }
```

## 实类举例
### 创建实例实现该抽象类
```
 /// <summary>
    /// 固定的Excel格式
    /// |Id|姓名|工号|科室|性别|人员类型|已删除|
    ///|---|---|---|---|---|---|
    ///|a3c6a1b9-8e54-46c9-a864-002b2cfc639e|姓名1|1234|ICU|女|护士|FALSE|
    ///startCellName 从姓名开始
    ///keyColumn为B列，即第2列
    /// </summary>
    public class ReadPersonInfoFromExcelService : ReadFromExcelService
    {
        private readonly UserManager<ApplicationUser> _userManager;

        public ReadPersonInfoFromExcelService(UserManager<ApplicationUser> userManager, string excelType, Stream fileStream, string startCellName = "B1", int mergeTitleRow = 1, int keyColumn = 1) : base(excelType, fileStream, startCellName, mergeTitleRow, keyColumn)
        {
            this._userManager = userManager;
        }

        /// <summary>
        /// 创建新用户
        /// </summary>
        /// <typeparam name="T">"Key"为各人员的Guid Id, Value为姓名、工号、科室、性别、人员类型、已删除</typeparam>
        /// <param name="messageFromExcel"></param>
        /// <returns></returns>
        protected override async Task<string> SaveDataFromExcelToDataBase<T>(Dictionary<string, T> messageFromExcel)
        {
            StringBuilder errorMsg = new StringBuilder(200);
            if(messageFromExcel != null && messageFromExcel.Count > 0)
            {
                var dataFromExcel = messageFromExcel as Dictionary<string, UserInfoFromExcelModel>;
                var newUserInfoModel = dataFromExcel.AsParallel().Where(a => a.Value.DeleteFlag != "TRUE").Select(a => new CreateUserViewModel { Id = Guid.Parse(a.Key), UserName = a.Value.EmployeeNo, FamilyName = a.Value.UserName[0].ToString(), FirstName = a.Value.UserName.Substring(1), Password = "123456", ConfirmPassword = "123456" }).ToList();
                foreach (var user in newUserInfoModel)
                {
                    await user.CreateUserAsync(_userManager);
                }
            }
            else
            {
                errorMsg.Append("读取Excel数据失败!");
            }
            return errorMsg.ToString();
        }
    }
```
### 调用实例操作
#### 前端界面
```
<form asp-action="ImportPersonInfo" method="post" enctype="multipart/form-data">
    <input type="file" name="file"  />
    <button type="submit" class="btn btn-primary">上传</button>
</form>

@if (ViewBag.ErrorMessage != null)
{
    <h5 class="text-danger">@ViewBag.ErrorMessage</h5>

}
```
#### 后台处理
```
  public async Task<IActionResult> ImportUserInfo(IFormFile file)
        {
            if (!IsValidFile(file))
            {
                ViewBag.ErrorMessage = "上传文件为空！";
                return View();
            }
            var fileStream = file.OpenReadStream();
            var result = await new ReadPersonInfoFromExcelService(this._userManager, file.ContentType, fileStream).ExcuteAsync<UserInfoFromExcelModel>();
            if (String.IsNullOrEmpty(result))
            {
                return RedirectToAction("ListUsers", "ManageUsers");

            }
            ViewBag.ErrorMessage = result;
            return View();
        }
         private bool IsValidFile(IFormFile file)
        {
            if (file != null)
            {
                return true;
            }
            return false;
        }
```