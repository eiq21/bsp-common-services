using Bsp.CommonServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http.Controllers;
using System.Web.Http.Filters;

namespace Bsp.CommonServices.Core
{
    public class LoggerAttributeApi : ActionFilterAttribute
    {
        public override void OnActionExecuting(HttpActionContext filterContext)
        {
            var controllerName = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            var actionName = filterContext.ActionDescriptor.ActionName;
            Logger.Info(string.Format("{0}: {1}.{2}", "Start", controllerName, actionName));
        }

        public override void OnActionExecuted(HttpActionExecutedContext filterContext)
        {
            var controllerName = filterContext.ActionContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            var actionName = filterContext.ActionContext.ActionDescriptor.ActionName;
            Logger.Info(string.Format("{0}: {1}.{2}", "Finish", controllerName, actionName));
        }
    }
}