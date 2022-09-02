using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Web;

namespace Bsp.CommonServices
{
    public class AppConfig
    {
        private static string applicationName;
        private static string applicationShortName;
        private static string applicationProductName;
        private static string applicationVersion;
        private static string applicationCopyright;

        /// <summary>
        /// Método para obtener o establecer el nombre completo de la aplicacion.
        /// </summary>
        public static string ApplicationName
        {
            set
            {
                applicationName = value;
            }
            get
            {
                if (applicationName != null && !string.IsNullOrEmpty(applicationName))
                {
                    return applicationName;
                }
                else
                {
                    return System.Configuration.ConfigurationManager.AppSettings[Constantes.KeyApplicationName].ToString() ?? Constantes.DefaultApplicationName;
                }
            }
        }


        /// <summary>
        /// Método para  obtener o establecer el nombre corto de la aplicacion (abreviatura).
        /// </summary>
        public static string ApplicationShortName
        {
            set
            {
                applicationShortName = value;
            }
            get
            {
                if (applicationShortName != null && !string.IsNullOrEmpty(applicationShortName))
                {
                    return applicationShortName;
                }
                else
                {
                    return System.Configuration.ConfigurationManager.AppSettings[Constantes.KeyApplicationShortName].ToString() ?? Constantes.DefaultApplicationShortName;
                }
            }
        }

        /// <summary>
        ///  Método para obtener o establecer la version de la aplicacion.
        /// </summary>
        public static string ApplicationVersion
        {
            set
            {
                applicationVersion = value;
            }
            get
            {
                if (applicationVersion != null && !string.IsNullOrEmpty(applicationVersion))
                {
                    return applicationVersion;
                }
                else
                {
                    var fileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AppConfig)).Location);
                    var version = string.Format("v{0}.{1}.{2}", fileVersionInfo.ProductMajorPart, fileVersionInfo.ProductMinorPart, fileVersionInfo.ProductBuildPart);

                    return version;
                }
            }
        }

        /// <summary>
        /// Método para obtener o establecer la licencia o derechos legales de la aplicacion.
        /// </summary>
        public static string ApplicationCopyright
        {
            set
            {
                applicationCopyright = value;
            }
            get
            {
                if (applicationCopyright != null && !string.IsNullOrEmpty(applicationCopyright))
                {
                    return applicationCopyright;
                }
                else
                {
                    var fileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AppConfig)).Location);

                    return fileVersionInfo.LegalCopyright;
                }
            }
        }

        /// <summary>
        /// Método para obtener o establecer el nombre de la aplicacion como producto.
        /// </summary>
        public static string ApplicationProductName
        {
            set
            {
                applicationProductName = value;
            }
            get
            {
                if (applicationProductName != null && !string.IsNullOrEmpty(applicationProductName))
                {
                    return applicationProductName;
                }
                else
                {
                    var fileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AppConfig)).Location);

                    return fileVersionInfo.ProductName;
                }
            }
        }
    }
}