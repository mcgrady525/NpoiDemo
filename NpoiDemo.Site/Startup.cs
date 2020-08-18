using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(NpoiDemo.Site.Startup))]
namespace NpoiDemo.Site
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
