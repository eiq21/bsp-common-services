using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.Entities
{
    public class ServidorEmailEntity
    {
        public string Host { get; set; }
        public int Puerto { get; set; }
        public bool HabitarSsl { get; set; }
        public bool CredencialesPorDefecto { get; set; }
        public CredencialEntity CredencialUsuario { get; set; }
    }
}