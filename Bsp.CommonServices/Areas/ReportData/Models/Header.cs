using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Bsp.CommonServices.Areas.ReportData.Models
{
    public class Header
    {
        public string TituloCabecera { get; set; }
        public int? IdTablaColumna { get; set; }
        public string IndexColumna { get; set; }
    }
}