using FileManager.Interfaces;
using Microsoft.Extensions.DependencyInjection;

namespace FileManager.IoC
{
    public static class IoC
    {
        public static void SetFileManagerDependencies(this IServiceCollection services)
        {
            services.AddTransient<ISearcher<string>, ReportA_Searcher>();
            services.AddTransient<ICopier<string, string>, ReportA_Copier>();
            services.AddTransient<IFileManager<string>, FileManager>();
        }
    }
}
