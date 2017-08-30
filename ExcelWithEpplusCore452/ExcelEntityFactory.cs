using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWithEpplusCore452
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelEntityFactory
    {
        #region single pattern
        private ExcelEntityFactory()
        {

        }

        // A private static instance of the same class
        private static readonly ExcelEntityFactory instance = null;

        static ExcelEntityFactory()
        {
            // create the instance only if the instance is null
            instance = new ExcelEntityFactory();
        }
        /// <summary>
        /// 获取实例
        /// </summary>
        /// <returns></returns>
        public static ExcelEntityFactory GetInstance()
        {
            // return the already existing instance
            return instance;
        }
        #endregion
        /// <summary>
        /// 构造ReadFromExcel
        /// </summary>
        /// <returns></returns>
        public IReadFromExcel CreateReadFromExcel()
        {
            IReadFromExcel result = null;
            result = new ReadFromExcel();
            return result;
        }
        /// <summary>
        /// CreateWriteToExcel
        /// </summary>
        /// <returns></returns>
        public IWriteToExcel CreateWriteToExcel()
        {
            IWriteToExcel result = null;
            result = new WriteToExcel();
            return result;
        }
        /// <summary>
        /// Excel2007ContentType
        /// </summary>
        public string Excel2007ContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }
    }
}
