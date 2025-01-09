using Infrastructure.Interfaces;
using Infrastructure.Launcher;
using Infrastructure.ViewModel;
using Microsoft.Extensions.DependencyInjection;

namespace Infrastructure.IoC
{
    public static class IoC
    {
        public static void SetInfrastructureDependencies(this IServiceCollection services)
        {
            services.AddSingleton<ILauncher, AppLauncher>();
            services.AddSingleton<IViewModel, MoscowReportVM>();
            services.AddSingleton<IProgress<object[]>, Progress<object[]>>();
        }
    }
}
