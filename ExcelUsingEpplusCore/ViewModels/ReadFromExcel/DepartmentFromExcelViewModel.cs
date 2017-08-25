using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUsingEpplusCore.ViewModels.ReadFromExcel
{
    public class DepartmentFromExcelViewModel
    {
        //[Display(Name = "科室名称")]
        //public string DepartmentName { get; set; }

        [Display(Name = "科室属性")]
        public string DepartmentType { get; set; }
    }
}
