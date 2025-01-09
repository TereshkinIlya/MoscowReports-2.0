using Core.Fabrics;
using Core.Interaces;
using Microsoft.Extensions.DependencyInjection;

namespace Core.IoC
{
    public static class IoC
    {
        public static void SetCoreDependencies(this IServiceCollection services)
        {
            services.AddTransient<ITable, TablesFactory>();
            services.AddTransient<IData, ContentFactory>();
            services.AddTransient<IHeader, HeaderFactory>();
        }
    }
}
