using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bsp.CommonServices.Common
{
    public class Configuracion
    {
        /// <summary>
        /// Obtiene el valor de la llave especificada en el archivo de configuracion.
        /// </summary>
        /// <param name="key">Llave de configuracion.</param>
        /// <returns>Retorna un string que representa el valor de la llave.</returns>
        public static string GetAppSetting(string key)
        {
            var value = ConfigurationManager.AppSettings[key];

            if (value != null)
            {
                return value.ToString();
            }
            return default(string);
        }
    }
}