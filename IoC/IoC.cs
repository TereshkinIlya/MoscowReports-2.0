using Microsoft.Extensions.DependencyInjection;
using FileManager.IoC;
using Infrastructure.IoC;
using ExcelConverter.IoC;
using ReportsHandler.IoC;
using TableHeadersHandler.IoC;
using Core.IoC;

namespace DI
{
    public static class IoC
    {
        public static void Register(this IServiceCollection services)
        {
            services.SetCoreDependencies();
            services.SetTableHeadersHandlerDependencies();
            services.SetFileManagerDependencies();
            services.SetInfrastructureDependencies();
            services.SetExcelConverterDependencies();
            services.SetReportsHandlerDependencies();
        }
    }

}
