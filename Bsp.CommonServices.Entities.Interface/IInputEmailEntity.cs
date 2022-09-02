using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bsp.CommonServices.Entities.Interface
{
    public interface IInputEmailEntity
    {
        IServidorEmailEntity Servidor { get; set; }
        IDireccionEmailEntity Remitente { get; set; }
        IEnumerable<IDireccionEmailEntity> Destinatarios { get; set; }
        string CadenaDestinatarios { get; set; }
        IEnumerable<IDireccionEmailEntity> DestinatariosCc { get; set; }
        string CadenaDestinatariosCc { get; set; }
        string Asunto { get; set; }
        IDictionary<string, string> ParametrosAsunto { get; set; }
        bool EsBodyHtml { get; set; }
        string Body { get; set; }
        string PathBody { get; set; }
        IDictionary<string, string> ParametrosBody { get; set; }
        string PathArchivoAdjunto { get; set; }
    }
}
