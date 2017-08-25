using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithEpplusCore
{
    /// <summary>
    /// 读取Excel文件内容
    /// </summary>
    public interface IReadFromExcel
    {
        /// <summary>
        /// 存在左侧与右侧分隔的格式的Excel，如A列是部门名称，B列开始是详细属性，如科室属性等，右侧标题行只有1行的情况。
        /// </summary>
        /// <typeparam name="T">与Excel标题栏匹配的模型</typeparam>
        /// <param name="TPropertyNameDisplayAttributeNameDic">Key为T的property.Name，Value是T的DisplayAttributeName</param>
        /// <param name="filePath">上传文件的路径</param>
        /// <param name="errorMsg">传出的错误</param>
        /// <param name="sheetName">工作表名称，默认为第一个</param>
        /// <param name="startCellName">内容开始的单元格，含标题</param>
        /// <param name="mergeTitleRow">标题的合并行数</param>
        /// <param name="keyColumn">Key值部门名称所在的列</param>
        /// <returns>如果只有一列有值，则Key为该列的值，且Value也为该列的值</returns>
        Dictionary<string, T> ExcelToEntityDictionary<T>(Dictionary<string, string> TPropertyNameDisplayAttributeNameDic, string filePath, out StringBuilder errorMsg, string sheetName = null, string startCellName = "A1", int mergeTitleRow = 1, int keyColumn = 1) where T : new();

        /// <summary>
        /// 存在左侧与右侧分隔的格式的Excel，如A列是部门名称，B列开始是详细属性，如科室属性等，右侧标题行只有1行的情况。
        /// </summary>
        /// <typeparam name="T">与Excel标题栏匹配的模型</typeparam>
        /// <param name="TPropertyNameDisplayAttributeNameDic">Key为T的property.Name，Value是T的DisplayAttributeName</param>
        /// <param name="fileStream">上传文件的Stream</param>
        /// <param name="errorMsg">传出的错误</param>
        /// <param name="sheetName">工作表名称，默认为第一个</param>
        /// <param name="startCellName">内容开始的单元格，含标题</param>
        /// <param name="keyColumn">Dictionary中Key所在的列，如A列则为1</param>
        /// <param name="mergeTitleRow">标题行的合并行数，有些标题行有两行</param>
        /// <returns></returns>
        //Dictionary<string, T> ExcelToEntityDictionary<T>(Dictionary<string, string> TPropertyNameDisplayAttributeNameDic, Stream fileStream, out StringBuilder errorMsg, string sheetName = null, string startCellName = "A1") where T : new();
        Dictionary<string, T> ExcelToEntityDictionary<T>(Dictionary<string, string> TPropertyNameDisplayAttributeNameDic, Stream fileStream, out StringBuilder errorMsg, string sheetName = null, string startCellName = "A1", int mergeTitleRow = 1, int keyColumn = 1) where T : new();

        /// <summary>
        ///
        /// 1、读取Excel文件并以此初始化一个工作簿(Workbook)；
        /// 2、从工作簿上获取一个工作表(Sheet)；默认为工作薄的第一个工作表；
        /// 3、遍历工作表所有的行(row)；第一行为标题行,生成一个包含行索引的Dictionary；
        /// 4、提供一个类属性名与Excel标题名相对应的Dictionary
        /// 5、遍历行的每一个单元格(cell)，根据一定的规律赋值给对象的属性。
        /// </summary>
        /// <typeparam name="T">与Excel标题栏匹配的模型</typeparam>
        /// <param name="cellHeard">"属性名DisplayName标题名"，T类型的属性名为Name，Value为[Display(Name="标题名")]</param>
        /// <param name="filePath">上传的Excel文件路径</param>
        /// <param name="errorMsg">传出的错误信息</param>
        /// <param name="sheetName">Excel表中的工作簿名称，默认为第1个</param>
        /// <param name="startCellName">Excel中数据行开始单元格,包括标题行</param>
        /// <param name="mergeTitleRow">标题行占用的行数</param>
        /// <returns></returns>
        List<T> ExcelToEntityList<T>(Dictionary<string, string> cellHeard, string filePath, out StringBuilder errorMsg, string sheetName = null, string startCellName = "A1", int mergeTitleRow = 1) where T : new();
    }
}
