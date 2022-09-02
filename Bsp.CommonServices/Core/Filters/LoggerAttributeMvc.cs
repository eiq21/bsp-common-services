using Bsp.CommonServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Bsp.CommonServices.Core
{
    public class LoggerAttributeMvc : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var controllerName = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            var actionName = filterContext.ActionDescriptor.ActionName;
            Logger.Info(string.Format("{0}: {1}.{2}", "Start", controllerName, actionName));
        }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {
            var controllerName = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            var actionName = filterContext.ActionDescriptor.ActionName;
            Logger.Info(string.Format("{0}: {1}.{2}", "Finish", controllerName, actionName));
        }
    }
}