using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithEpplusCore
{
    /// <summary>
    /// 导出Excel
    /// </summary>
    public interface IWriteToExcel
    {
        /// <summary>
        /// 将Dictionary的值写入到Excel中，如果showKey为true，将外层的key写入在每一列的第一行
        /// 内层Dictionary从第开始列开始，key为第一行内容，value为第二行内容。形如
        /// |标识|(姓名）第一Key|（部门）第二Key|
        /// |Key|张三|人事部|
        /// |标识|(姓名）第一Key|（部门）第二Key|
        /// |Key|李四|技术部|
        /// </summary>
        /// <param name="data">需写入到Excel中的数据</param>
        /// <param name="showKey">外层Dictionary的key是否写入到excel，默认不写</param>
        /// <returns></returns>
        byte[] ExportDictionaryToExcel(Dictionary<string, Dictionary<string, string>> data, bool showKey = false);
        /// <summary>
        /// 对应的模版为第一行为大标题，第1列为各关键列，第二行为关键列对应的各属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceData">key为Excel模版中的keyCell对应列的内容，如A1列中为"部门1，部门2，部门3","部门1, T"，可将部门1等内容写在属性的Display(Name="部门1")上</param>
        /// <param name="templateFilePath">Excel模版路径</param>
        /// <param name="Message">输出错误信息</param>
        /// <param name="dataRangeStartCell">数据区域开始的单元格，如第1行为标题，A列为关键字，数据区域从B2开始写</param>
        /// <param name="keyCell">如A1列中为"部门1，部门2，部门3"，标题为“姓名、年龄、生日”等</param>
        /// <param name="changeTitleCell">需更改的标题所在单元格</param>
        /// <param name="changeContent">将标题所在单元格替换为该内容</param>
        /// <param name="columnData">Key为对应单元格，Value为T的属性名，将属性值写入到对应单元格内，如//"C5,Name" "D5, Age" Name与Age均为T中的属性，表示将Name中的值写入到C5</param>
        /// <param name="sheetName">工作表名称，默认为第一个</param>
        /// <returns>写入成功，返回True，否则返回false</returns>
        bool ExportEntityToExcelFile<T>(Dictionary<string, T> sourceData, string templateFilePath, out StringBuilder Message, string dataRangeStartCell, string keyCell, string changeTitleCell, string changeContent, Dictionary<string, string> columnData, string sheetName) where T : class;
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">类型T</typeparam>
        /// <param name="data">传入数据</param>
        /// <param name="heading">第一行标题名称</param>
        /// <param name="isShowSlNo">在第一列显示序号</param>
        /// <returns></returns>
        byte[] ExportListToExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false);

        /// <summary>
        /// Excel2007的类型
        /// </summary>
        string ExcelContentType { get; }

    }
}
