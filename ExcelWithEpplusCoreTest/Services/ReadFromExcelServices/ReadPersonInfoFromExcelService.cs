using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWithEpplusCoreTest.Services.ReadFromExcelServices
{
    public class ReadPersonInfoFromExcelService : ReadFromExcelService
    {
        public ReadPersonInfoFromExcelService(string excelType, Stream fileStream, string startCellName, int keyColumn) : base(excelType, fileStream, startCellName, keyColumn)
        {
        }

        protected override Task<string> SaveDataFromExcelToDataBase<T>(Dictionary<string, T> messageFromExcel)
        {
            throw new NotImplementedException();
        }
    }
}
