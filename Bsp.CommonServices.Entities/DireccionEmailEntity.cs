using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.Entities
{
    public class DireccionEmailEntity : IDireccionEmailEntity
    {
        public string Direccion { get; set; }
        public string Nombre { get; set; }
    }
}