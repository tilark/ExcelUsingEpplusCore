using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWithEpplusCore452.ExcelServices
{
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

        public static ExcelEntityFactory GetInstance()
        {
            // return the already existing instance
            return instance;
        }
        #endregion

        public IReadFromExcel CreateReadFromExcel()
        {
            IReadFromExcel result = null;
            result = new ReadFromExcel();
            return result;
        }

        public IWriteToExcel CreateWriteToExcel()
        {
            IWriteToExcel result = null;
            result = new WriteToExcel();
            return result;
        }
    }
}
