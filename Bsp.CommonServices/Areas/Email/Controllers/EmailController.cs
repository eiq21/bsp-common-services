using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web.Http;
using Bsp.CommonServices.Entities;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.Areas.Email.Controllers
{
    public class EmailController : ApiController
    {
        BusinessLogic.EmailBL EmailBL;

        public EmailController(): this(new BusinessLogic.EmailBL())
        {
        }

        public EmailController(Bsp.CommonServices.BusinessLogic.EmailBL emailBL)
        {
            EmailBL = emailBL;
        }

        public async Task<string> PostEnviar(InputEmailEntity inputEmailEntity)
        {
            return await EmailBL.EnviarEmailAsync(inputEmailEntity);
        }
    }
}
