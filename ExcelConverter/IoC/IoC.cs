using ExcelConverter.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelConverter.IoC
{
    public static class IoC
    {
        public static void SetExcelConverterDependencies(this IServiceCollection services)
        {
            services.AddTransient<InteropExcelApp>();
            services.AddTransient(provider =>
                new Excel.Application
                {
                    DisplayAlerts = false,
                    Visible = false
                }
            );
            services.AddTransient<IExcelConverter<string>, InteropExcelConverter>();
        }
    }
}
