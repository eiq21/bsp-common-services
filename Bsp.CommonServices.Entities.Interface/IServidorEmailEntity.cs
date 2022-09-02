using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bsp.CommonServices.Entities.Interface
{
    public interface IServidorEmailEntity
    {
        string Host { get; set; }
        int Puerto { get; set; }
        bool HabitarSsl { get; set; }
        bool CredencialesPorDefecto { get; set; }
        ICredencialEntity CredencialUsuario { get; set; }
    }
}