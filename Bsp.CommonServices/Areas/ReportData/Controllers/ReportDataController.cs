using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using Bsp.CommonServices.Areas.ReportData.ExcelHelps;
using Bsp.CommonServices.Areas.ReportData.Helps;
using Bsp.CommonServices.Areas.ReportData.Models;
using Bsp.CommonServices.BusinessLogic;
using Bsp.CommonServices.Entities;
using Newtonsoft.Json;
using NPOI.SS.UserModel;

namespace Bsp.CommonServices.Areas.ReportData.Controllers
{
    [System.Web.Http.Route("api/reportdata")]
    public class ReportDataController : ApiController
    {


        [System.Web.Http.HttpPost]
        public async Task<ResultRequest> PostAsync([FromBody]DataResponse data)
        {
            ResultRequest result = null;

            try
            {
                Task<int> tGenerateExcel = Task.Factory.StartNew(() => new GenerateExcelFile().ExportaExcelDetalle("true", data.IdReporte.ToString(),
                   data.FechaReporteNumerico.ToString(),
                   data.FechaReporte,
                   "",
                   "",
                  data.DescripcionReporte,
                   data));
                int r = await tGenerateExcel;
                if (r > 0)
                {
                    await SendEmail(data.InputEmailEntity);

                    result = new ResultRequest() { IsError = false, PathFile = "" };
                }
                else
                    result = new ResultRequest() { IsError = true, PathFile = null };

            }
            catch (Exception)
            {
                result = new ResultRequest() { IsError = true, PathFile = null };
            }

            return result;

        }

        private static async Task SendEmail(InputEmailEntity mail)
        {
            try
            {
                
                await new EmailBL().EnviarEmailAsync(mail);
            }
            catch (Exception)
            {

            }
           
        }
    }
}
