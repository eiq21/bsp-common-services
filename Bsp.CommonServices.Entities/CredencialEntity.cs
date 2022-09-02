using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.Entities
{
    public class CredencialEntity : ICredencialEntity
    {
        public string NombreUsuario { get; set; }
        public string Password { get; set; }
    }
}