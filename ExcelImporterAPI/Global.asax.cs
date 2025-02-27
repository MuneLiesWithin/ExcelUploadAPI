using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Threading.Tasks;
using System.Threading;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace ExcelImporterAPI
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            GlobalConfiguration.Configuration.MessageHandlers.Add(new CorsHandler());
        }

        public class CorsHandler : DelegatingHandler
        {
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                if (request.Headers.Contains("Origin"))
                {
                    if (request.Method == HttpMethod.Options)
                    {
                        var response = new HttpResponseMessage(HttpStatusCode.OK);
                        response.Headers.Add("Access-Control-Allow-Origin", "*");
                        response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
                        response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
                        response.Headers.Add("Access-Control-Allow-Credentials", "true");

                        return Task.FromResult(response);
                    }
                    else
                    {
                        var response = base.SendAsync(request, cancellationToken).Result;
                        response.Headers.Add("Access-Control-Allow-Origin", "*");
                        response.Headers.Add("Access-Control-Allow-Credentials", "true");

                        return Task.FromResult(response);
                    }
                }

                return base.SendAsync(request, cancellationToken);
            }
        }
    }
}
