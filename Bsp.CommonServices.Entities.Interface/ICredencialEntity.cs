using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bsp.CommonServices.Entities.Interface
{
    public interface ICredencialEntity
    {
        string NombreUsuario { get; set; }
        string Password { get; set; }
    }
}
