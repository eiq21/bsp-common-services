using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Diagnostics;
/*
 * =============== CREATION ===================================
 * Author: Vector Software Factory Peru | http://www.vectoritcgroup.com/
 * Date: 2018-02-09
 * Request: CEEFFC Etapa 2 
 * Description: CEEFFC Etapa 2
 * =============== MODIFICATION 0000 (MOD.0000) ===============
 * Author:
 * Date:
 * Request:
 * Description:
 * ============================================================
 */
namespace Bsp.CommonServices
{
    /// <summary>
    /// Permite registrar eventos en un archivo fisico log.
    /// </summary>
    public class Logger : IDisposable
    {
        private static NLog.Logger logger = LogManager.GetCurrentClassLogger();
        private string FullNameCaller;

        /// <summary>
        /// Constructor de la clase que permite registrar un evento que indica el inicio de la ejecucion de una 
        /// funcionalidad. La finalizacion se realiza automaticamente cuando se crea el objeto Logger con la 
        /// sentencia "using".
        /// </summary>
        /// <param name="obj">Objeto para indentificar el origen de la funcionalidad.</param>
        public Logger(object obj)
        {
            StackTrace st = new StackTrace();
            FullNameCaller = string.Format("{0}>{1}", obj.GetType().FullName, st.GetFrame(1).GetMethod().Name);
            logger.Info("Start: " + FullNameCaller);
            st = null;
        }

        /// <summary>
        /// Permite registrar un evento del tipo informativo.
        /// </summary>
        /// <param name="message">Mensaje a registrar.</param>
        public static void Info(string message)
        {
            logger.Info(message);
        }

        /// <summary>
        /// Permite registrar un evento del tipo informativo.
        /// </summary>
        /// <param name="exception">Objecto de excepcion a registrar.</param>
        public static void Info(Exception exception)
        {
            StackTrace st = new StackTrace();
            var error = string.Format("{0}>{1}. Error: {2}", st.GetFrame(1).GetMethod().ReflectedType.FullName, st.GetFrame(1).GetMethod().Name, exception.Message);
            logger.Info(error);
            st = null;
        }

        /// <summary>
        /// Permite registrar un evento del tipo advertencia.
        /// </summary>
        /// <param name="message">Mensaje a registrar.</param>
        public static void Warn(string message)
        {
            logger.Warn(message);
        }

        /// <summary>
        /// Permite registrar un evento del tipo advertencia.
        /// </summary>
        /// <param name="exception">Objeto de excepcion a registrar.</param>
        public static void Warn(Exception exception)
        {
            StackTrace st = new StackTrace();
            var error = string.Format("{0}>{1}. Error: {2}", st.GetFrame(1).GetMethod().ReflectedType.FullName, st.GetFrame(1).GetMethod().Name, exception.Message);
            logger.Warn(error);
            st = null;
        }

        /// <summary>
        /// Permite registrar un evento del tipo error.
        /// </summary>
        /// <param name="message">Mensaje a registrar.</param>
        public static void Error(string message)
        {
            logger.Error(message);
        }

        /// <summary>
        /// Permite registrar un evento del tipo error.
        /// </summary>
        /// <param name="exception">Objeto de excepcion a registrar.</param>
        public static void Error(Exception exception)
        {
            StackTrace st = new StackTrace();
            var error = string.Format("{0}>{1}. Error: {2}", st.GetFrame(1).GetMethod().ReflectedType.FullName, st.GetFrame(1).GetMethod().Name, exception.Message);
            logger.Error(error);
            st = null;
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    logger.Info("Finish: " + FullNameCaller);
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~Logger() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        /// <summary>
        /// Permite registrar un evento que indica el fin de la ejecucion de una funcionalidad. 
        /// La finalizacion se realiza automaticamente cuando se crea el objeto Logger con la sentencia "using".
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}