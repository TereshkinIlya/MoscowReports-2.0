using System.Windows;
using DI;
using Microsoft.Extensions.DependencyInjection;

namespace MoscowReports_2._0
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private readonly ServiceProvider _serviceProvider;
        public App()
        {
            IServiceCollection services = new ServiceCollection();
            services.Register();
            services.AddSingleton<MainWindow>();
            _serviceProvider = services.BuildServiceProvider();
        }
        private void OnStartup (object sender, StartupEventArgs e)
        {
            Window? mainWindow = _serviceProvider.GetService<MainWindow>();
            mainWindow?.Show();
        }


    }
}
