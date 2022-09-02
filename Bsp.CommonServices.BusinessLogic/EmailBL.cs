using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Bsp.CommonServices.Common;
using Bsp.CommonServices.Entities;
using Bsp.CommonServices.Entities.Interface;

namespace Bsp.CommonServices.BusinessLogic
{
    public class EmailBL
    {
        public Task<string> EnviarEmailAsync(InputEmailEntity inputEmailEntity)
        {
            var taskInvoke = Task<string>.Factory.StartNew(() =>
            {
                try
                {
                    ValidarDatos(inputEmailEntity);

                    var mailMessage = new MailMessage();

                    // Destinatarios
                    mailMessage.To.Add(ObtenerDestinatarios(inputEmailEntity.Destinatarios, inputEmailEntity.CadenaDestinatarios));

                    // Destinatarios Cc
                    if ((inputEmailEntity.DestinatariosCc != null && inputEmailEntity.DestinatariosCc.Any()) || !string.IsNullOrEmpty(inputEmailEntity.CadenaDestinatariosCc))
                    {
                        mailMessage.CC.Add(ObtenerDestinatarios(inputEmailEntity.DestinatariosCc, inputEmailEntity.CadenaDestinatariosCc));
                    }

                    // Asunto
                    mailMessage.Subject = ObtenerAsunto(inputEmailEntity.Asunto, inputEmailEntity.ParametrosAsunto);
                    mailMessage.SubjectEncoding = Encoding.UTF8;

                    // Body
                    mailMessage.Body = ObtenerBody(inputEmailEntity.Body, inputEmailEntity.ParametrosBody, inputEmailEntity.PathBody);
                    mailMessage.BodyEncoding = Encoding.UTF8;
                    mailMessage.IsBodyHtml = inputEmailEntity.EsBodyHtml;

                    // Remitente
                    mailMessage.From = new MailAddress(inputEmailEntity.Remitente.Direccion, inputEmailEntity.Remitente.Nombre);

                    // Archivo 
                    if (!string.IsNullOrEmpty(inputEmailEntity.PathArchivoAdjunto))
                    {
                        mailMessage.Attachments.Add(ObtenerArchivoAdjunto(inputEmailEntity.PathArchivoAdjunto));
                    }

                    // Servidor
                    SmtpClient cliente = new SmtpClient();
                    cliente.UseDefaultCredentials = inputEmailEntity.Servidor.CredencialesPorDefecto;
                    cliente.Credentials = new System.Net.NetworkCredential(inputEmailEntity.Servidor.CredencialUsuario.NombreUsuario, inputEmailEntity.Servidor.CredencialUsuario.Password);
                    cliente.Host = inputEmailEntity.Servidor.Host;
                    cliente.Port = inputEmailEntity.Servidor.Puerto;
                    cliente.EnableSsl = inputEmailEntity.Servidor.HabitarSsl;

                    // Enviar mensaje
                    cliente.Send(mailMessage);
                    return "OK";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            });

            return taskInvoke;
        }

        string ObtenerDestinatarios(IEnumerable<IDireccionEmailEntity> destinatarios, string cadenaDestinatarios)
        {
            var emails = string.Empty;

            if (destinatarios != null && destinatarios.Any())
            {
                foreach (var item in destinatarios)
                {
                    if (!string.IsNullOrEmpty(item.Direccion))
                    {
                        if (!string.IsNullOrEmpty(item.Nombre))
                        {
                            emails += string.Format("{0}<{1}>,", item.Nombre, item.Direccion);
                        }
                        else
                        {
                            emails += string.Format("{0},", item.Direccion);
                        }
                    }
                }
            }
            else
            {
                var caracter = Convert.ToChar(Configuracion.GetAppSetting(Constantes.KeySeparadorEmails));
                var lista = cadenaDestinatarios.Split(caracter);

                foreach (var item in lista)
                {
                    emails += string.Format("{0},", item);
                }
            }

            emails = emails.TrimEnd(',');

            return emails;
        }

        string ObtenerAsunto(string asunto, IDictionary<string, string> parametrosAsunto)
        {
            foreach (var item in parametrosAsunto.Keys)
            {
                asunto = asunto.Replace("{{" + item + "}}", parametrosAsunto[item]);
            }
            return asunto;
        }

        string ObtenerBody(string body, IDictionary<string, string> parametrosBody, string pathBody)
        {
            if (!string.IsNullOrEmpty(pathBody))
            {
                StreamReader sr = new StreamReader(pathBody);
                body = sr.ReadToEnd();
            }

            foreach (var item in parametrosBody.Keys)
            {
                body = body.Replace("{{" + item + "}}", parametrosBody[item]);
            }
            return body;
        }

        Attachment ObtenerArchivoAdjunto(string pathArchivoAdjunto)
        {
            var adjunto = new Attachment(pathArchivoAdjunto, MediaTypeNames.Application.Octet);
            var disposicion = adjunto.ContentDisposition;
            disposicion.CreationDate = File.GetCreationTime(pathArchivoAdjunto);
            disposicion.ModificationDate = File.GetLastWriteTime(pathArchivoAdjunto);
            disposicion.ReadDate = File.GetLastAccessTime(pathArchivoAdjunto);
            return adjunto;
        }

        void ValidarDatos(InputEmailEntity inputEmailEntity)
        {
            if (inputEmailEntity == null)
            {
                throw new Exception("Debe especificar los parámetros del mensaje.");
            }
            if ((inputEmailEntity.Destinatarios == null || !inputEmailEntity.Destinatarios.Any()) && string.IsNullOrEmpty(inputEmailEntity.CadenaDestinatarios))
            {
                throw new Exception("Debe especificar al menos un destinatario válido.");
            }
            if (inputEmailEntity.Destinatarios != null && inputEmailEntity.Destinatarios.Any() &&
                inputEmailEntity.Destinatarios.Where(x => string.IsNullOrEmpty(x.Direccion)).ToList().Count == inputEmailEntity.Destinatarios.ToList().Count())
            {
                throw new Exception("Debe especificar al menos un destinatario válido.");
            }
            if (inputEmailEntity.Remitente == null || string.IsNullOrEmpty(inputEmailEntity.Remitente.Direccion))
            {
                throw new Exception("Debe especificar un remitente válido.");
            }
            if (inputEmailEntity.Servidor == null)
            {
                throw new Exception("Debe especificar los parámetros del servidor de correo electrónico.");
            }
            if (string.IsNullOrEmpty(inputEmailEntity.Servidor.Host))
            {
                throw new Exception("Debe especificar el host del servidor de correo electrónico.");
            }
            if (inputEmailEntity.Servidor.Puerto == 0)
            {
                throw new Exception("Debe especificar el puerto del host del servidor de correo electrónico.");
            }
            if (inputEmailEntity.Servidor.CredencialUsuario == null
                || string.IsNullOrEmpty(inputEmailEntity.Servidor.CredencialUsuario.NombreUsuario)
                || string.IsNullOrEmpty(inputEmailEntity.Servidor.CredencialUsuario.Password))
            {
                throw new Exception("Debe especificar las credenciales de usuario para iniciar sesión en el servidor de correo electrónico.");
            }
            if (!string.IsNullOrEmpty(inputEmailEntity.PathBody) && !System.IO.File.Exists(inputEmailEntity.PathBody))
            {
                throw new Exception("La ruta de la plantilla del body no es válida o no es accesible.");
            }
            if (!string.IsNullOrEmpty(inputEmailEntity.PathArchivoAdjunto) && !System.IO.File.Exists(inputEmailEntity.PathArchivoAdjunto))
            {
                throw new Exception("La ruta del archivo adjunto no es válida o no es accesible.");
            }
        }
    }
}