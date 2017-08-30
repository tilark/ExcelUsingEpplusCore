using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWithEpplusCore452.ViewModels
{
    public class TestA
    {
        [Display(Name = ("名称"))]
        public virtual string Name { get; set; }

        [Display(Name = ("排序号码"))]
        public virtual long SortNumber { get; set; }

        [Display(Name = ("备注"))]
        public virtual string Remarks { get; set; }
    }
}
