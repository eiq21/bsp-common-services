using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Bsp.CommonServices.Entities;

namespace Bsp.CommonServices.Areas.ReportData.Models
{
    public class DataResponse
    {
        public InputEmailEntity InputEmailEntity { get; set; }
        public string CodigoEmpresa { get; set; }
        public string NroDocumento { get; set; }
        public string DescripcionReporte { get; set; }
        public string FechaReporte { get; set; }
        public int FechaReporteNumerico { get; set; }
        public ICollection<Header> Header { get; set; }
        public ICollection<Details> Details { get; set; }
        public int? IdReporte { get; set; }
        public string Ruta { get; set; }
    }
}