using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(excelFunc2.Startup))]
namespace excelFunc2
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
