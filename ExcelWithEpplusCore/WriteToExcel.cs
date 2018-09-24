using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.ComponentModel;
using System.Reflection;

namespace ExcelWithEpplusCore
{
    /// <summary>
    /// 
    /// </summary>
    public class WriteToExcel : IWriteToExcel
    {

        #region Excel2007版本类型
        /// <summary>
        /// 
        /// </summary>
        public string ExcelContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }

        #endregion
        #region 将数据写入到指定Excel模版中
        /// <summary>
        /// 对应的模版为第一行为大标题，第1列为各关键列，第二行为关键列对应的各属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceData">"标题,T"key为Excel模版中的keyCell对应列的内容，如A1列中为"部门1，部门2，部门3","部门1, T"，可将部门1等内容写在属性的Display(Name="部门1")上</param>
        /// <param name="templateFilePath">Excel模版路径</param>
        /// <param name="Message">输出错误信息</param>
        /// <param name="dataRangeStartCell">数据区域开始的单元格，如第1行为标题，A列为关键字，数据区域从B2开始写</param>
        /// <param name="keyCell">如A1列中为"部门1，部门2，部门3"，标题为“姓名、年龄、生日”等</param>
        /// <param name="changeTitleCell">需更改的标题所在单元格</param>
        /// <param name="changeContent">将标题所在单元格替换为该内容</param>
        /// <param name="columnData">Key为对应单元格，Value为T的属性名，将属性值写入到对应单元格内，如//"C5,Name" "D5, Age" Name与Age均为T中的属性，表示将Name中的值写入到C5</param>
        /// <param name="sheetName">工作表名称，默认为第一个</param>
        /// <returns>写入成功，返回True，否则返回false</returns>
        public bool ExportEntityToExcelFile<T>(Dictionary<string, T> sourceData, string templateFilePath, out StringBuilder Message, string dataRangeStartCell, string keyCell, string changeTitleCell, string changeContent, Dictionary<string, string> columnData, string sheetName) where T : class
        {
            bool result = false;
            var message = new StringBuilder(100);
            try
            {
                FileInfo existingFile = new FileInfo(templateFilePath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    if (!String.IsNullOrEmpty(sheetName))
                    {
                        //如果没有获取该名称的sheet，获取第一个
                        worksheet = package.Workbook.Worksheets[sheetName];
                    }
                    if (worksheet.Dimension == null)
                    {
                        message.Append("EmptyError:File is Empty");
                        Message = message;
                        return result;
                    }
                    else
                    {
                        var keyRowStart = GetRowIndex(keyCell);
                        var keyColumn = GetColumnName(keyCell)[0] - 'A' + 1;
                        var dataRowStart = GetRowIndex(dataRangeStartCell);
                        var dataColumnStart = GetColumnName(dataRangeStartCell);
                        //需获得keyColumn中的Rows行
                        var keyHeader = new Dictionary<int, string>();
                        int colEnd = worksheet.Dimension.End.Column;
                        int rowEnd = worksheet.Dimension.End.Row;
                        for (int i = (int)keyRowStart; i <= rowEnd; i++)
                        {
                            if (worksheet.Cells[i, keyColumn].Value != null)
                            {
                                keyHeader[i] = worksheet.Cells[i, keyColumn].Value.ToString();

                            }
                        }
                        //
                        //更改标题栏的年与月
                        worksheet.Cells[changeTitleCell].Value = changeContent;

                        //将sourceData转换成Dictionary<row, T>形式
                        var destData = ConvertDataSourceToRowDataDictionary<T>(sourceData, keyHeader);
                        //制作原值-固定表和制作原值-无形表
                        for (int dataRow = (int)dataRowStart; dataRow <= rowEnd; dataRow++)
                        {
                            if (worksheet.Cells[dataRow, keyColumn].Value == null)
                            {
                                break;
                            }
                            //
                            //  name="columnData">Key为对应单元格，Value为T的属性名，将属性值写入到对应单元格内，如//"C5,Name" "D5, Age" Name与Age均为T中的属性，表示将Name中的值写入到C5
                            foreach (var columnName in columnData)
                            {
                                var cellName = columnName.Key + dataRow.ToString();
                                var propertiInfo = typeof(T).GetProperty(columnName.Value);
                                var pointData = propertiInfo.GetValue(destData[dataRow]);
                                worksheet.Cells[cellName].Value = pointData;
                            }

                        }
                    }
                    package.Save();
                    result = true;
                }
            }
            catch (Exception)
            {

                throw;
            }
            Message = message;
            return result;
        }


        #endregion

        #region 将数据写入到byte[]中
        /// <summary>
        /// 将数据写入到byte[]中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="heading"></param>
        /// <param name="isShowSlNo"></param>
        /// <returns></returns>
        public byte[] ExportListToExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false)
        {
            //return ExportExcel(ListToDataTable<T>(data), heading, isShowSlNo, ColumnsToTake);
            return WriteListDataToExcel(data, heading, isShowSlNo);

        }
        private byte[] WriteListDataToExcel<T>(List<T> data, List<string> heading, bool isShowSlNo)
        {
            byte[] results = null;
            using (ExcelPackage package = new ExcelPackage())
            {
                //var worksheetName = heading.Count > 0 ? heading[0] : String.Empty;
                var worksheetName = "OutPutExcel";

                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(string.Format("{0} Data", worksheetName));

                int startRowIndex = 1;
                int startColumnIndex = 1;
                //在第一列加入序号
                if (isShowSlNo)
                {
                    //加入序号                   
                    workSheet.Cells[startRowIndex, 1].Value = "序号";

                    for (int slNo = 0; slNo < data.Count; slNo++)
                    {
                        workSheet.Cells[slNo + 2, 1].Value = slNo + 1;
                    }
                    //起始列改为从第2行开始
                    startColumnIndex = 2;
                }

                //在第一行加入Heading
                if (heading != null && heading.Count > 0)
                {
                    for (int columnIndex = 0; columnIndex < heading.Count; columnIndex++)
                    {
                        workSheet.Cells[1, columnIndex + startColumnIndex].Value = heading[columnIndex];
                    }
                    //起始行改为从第2行开始
                    startRowIndex = 2;
                    workSheet.Cells[startRowIndex, startColumnIndex].LoadFromCollection(data, false);
                }
                else
                {
                    workSheet.Cells[startRowIndex, startColumnIndex].LoadFromCollection(data, true);
                }

                results = package.GetAsByteArray();
            }

            return results;
        }


        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sourceData"></param>
        /// <param name="keyHeader"></param>
        /// <returns></returns>
        private Dictionary<int, T> ConvertDataSourceToRowDataDictionary<T>(Dictionary<string, T> sourceData, Dictionary<int, string> keyHeader) where T : class
        {
            var result = new Dictionary<int, T>();
            foreach (var temp in keyHeader)
            {
                try
                {
                    if (sourceData.Keys.Contains(temp.Value))
                    {
                        var colNumber = sourceData[temp.Value];
                        result.Add(temp.Key, colNumber);
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            return result;
        }

        internal uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
        // Given a cell name, parses the specified cell to get the column name.
        internal string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="showKey"></param>
        /// <returns></returns>
        public byte[] ExportDictionaryToExcel(Dictionary<string, Dictionary<string, string>> data, bool showKey = false)
        {
            byte[] results = null;
            using (ExcelPackage package = new ExcelPackage())
            {
                //var worksheetName = heading.Count > 0 ? heading[0] : String.Empty;
                var worksheetName = "OutPutExcel";

                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(string.Format("{0} Data", worksheetName));

                int rowIndex = 1;
                int columnIndex = 1;
                //从第一行第一列开始写数
                foreach(var firstDic in data)
                {
                    // 如果showKey = true，则将第一个key写入到第一列

                    if (showKey)
                    {
                        workSheet.Cells[rowIndex, columnIndex].Value = "标识";
                        workSheet.Cells[rowIndex+1, columnIndex].Value = firstDic.Key;
                        columnIndex++;
                    }
                    foreach(var secondDic in firstDic.Value)
                    {
                        // 开始将data写入
                        workSheet.Cells[rowIndex, columnIndex].Value = secondDic.Key;
                        workSheet.Cells[rowIndex + 1, columnIndex].Value = secondDic.Value;
                        columnIndex++;
                    }
                    columnIndex = 1;
                    rowIndex += 2;
                }             

                results = package.GetAsByteArray();
            }

            return results;
        }
    }
}
