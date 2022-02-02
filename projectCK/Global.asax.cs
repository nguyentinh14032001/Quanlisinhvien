using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace projectCK
{
    public class MvcApplication : System.Web.HttpApplication
    {
       

        protected void Application_Start()
        {
            var config = GlobalConfiguration.Configuration;
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
         config.Formatters.JsonFormatter.SerializerSettings.ReferenceLoopHandling
          = Newtonsoft.Json.ReferenceLoopHandling.Ignore;
         config.Formatters.JsonFormatter.SerializerSettings.PreserveReferencesHandling
     = Newtonsoft.Json.PreserveReferencesHandling.Objects;

        }
    }
}
