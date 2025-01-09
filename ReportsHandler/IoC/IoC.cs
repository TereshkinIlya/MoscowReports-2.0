using ClosedXML.Excel;
using Core.Abstracts;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;
using ReportsHandler.Repositories;
using ReportsHandler.Services.Readers;
using ReportsHandler.Services.Writers;

namespace ReportsHandler.IoC
{
    public static class IoC
    {
        public static void SetReportsHandlerDependencies(this IServiceCollection services)
        {

            services.AddSingleton<IReportHandler<string>, ReportsHandler>();
            services.AddSingleton<IExcel<XLWorkbook>, ExcelBehaviour>();
            services.AddKeyedSingleton<Repository<Queue<Content>, Content>, SmallReportsRepo>("smallsRepo");
            services.AddKeyedSingleton<Repository<Queue<Content>, Content>, BigReportsRepo>("bigsRepo");
            services.AddKeyedSingleton<Repository<Queue<Content>, Content>, PiktsInfoRepo>("piktsRepo");

            services.AddTransient<IExcelReader<XLWorkbook>, BigReportReader>();
            services.AddTransient<IExcelReader<XLWorkbook>, SmallReportReader>();
            services.AddTransient<IExcelReader<XLWorkbook>, PiktsReportReader>();
            services.AddTransient<IExcelWriter<XLWorkbook>, BigReportWriter>();
            services.AddTransient<IExcelWriter<XLWorkbook>, SmallReportWriter>();
            services.AddTransient<IExcelWriter<XLWorkbook>, PiktsInfoWriter>();
        }
    }
}
