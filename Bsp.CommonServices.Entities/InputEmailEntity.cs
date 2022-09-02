using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.Entities
{
    public class InputEmailEntity
    {
        public ServidorEmailEntity Servidor { get; set; }
        public DireccionEmailEntity Remitente { get; set; }
        public IEnumerable<DireccionEmailEntity> Destinatarios { get; set; }
        public string CadenaDestinatarios { get; set; }
        public IEnumerable<DireccionEmailEntity> DestinatariosCc { get; set; }
        public string CadenaDestinatariosCc { get; set; }
        public string Asunto { get; set; }
        public IDictionary<string, string> ParametrosAsunto { get; set; }
        public bool EsBodyHtml { get; set; }
        public string Body { get; set; }
        public string PathBody { get; set; }
        public IDictionary<string, string> ParametrosBody { get; set; }
        public string PathArchivoAdjunto { get; set; }
    }
}