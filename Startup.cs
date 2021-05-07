using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(QLTV.Startup))]
namespace QLTV
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
