using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using TableHeadersHandler.Interfaces;
using TablesHandler;

namespace TableHeadersHandler.IoC
{
    public static class IoC
    {
        public static void SetTableHeadersHandlerDependencies(this IServiceCollection services)
        {
            services.AddTransient<IHeadline<XLWorkbook>, HeaderInitializer>();
            services.AddTransient<ISpreadsheet, TableHeadline>();
        }
    }
}
