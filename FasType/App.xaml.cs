using FasType.LLKeyboardListener;
using FasType.Windows;
using FasType.Pages;
using FasType.Services;
using FasType.Storage;
using FasType.ViewModels;
using IWshRuntimeLibrary;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
//using Microsoft.Extensions.Options;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace FasType
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static new App Current => (App)Application.Current;
        public IServiceProvider ServiceProvider { get; private set; }
        public IConfiguration Configuration { get; private set; }

        public App()
        {
            Configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            ServiceProvider = serviceCollection.BuildServiceProvider();
        }

        private void ConfigureServices(ServiceCollection services)
        {
            Log.Logger = new LoggerConfiguration()
                //.ReadFrom.Configuration(Configuration)
                .MinimumLevel.Verbose()
                .WriteTo.Debug()
                .CreateLogger();

            services.AddSingleton(Configuration);
            services.AddSingleton<MainWindow>();
            services.AddTransient<MainWindowViewModel>();

            services.AddSingleton<IDataStorage, FileDataStorage>();
            //services.AddTransient<IKeyboardListenerHandler, KeyboardListenerHandler>();
            
            services.AddTransient<ToolWindow>();
            services.AddTransient<SimpleAbbreviationPage>();
            services.AddTransient<SimpleAbbreviationViewModel>();

            services.AddTransient<SeeAllWindow>();
            services.AddTransient<SeeAllViewModel>();
        }

        private void OnStartup(object sender, StartupEventArgs args)
        {
            MainWindow = ServiceProvider.GetRequiredService<MainWindow>();
            MainWindow.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            FasType.Properties.Settings.Default.Save();
            base.OnExit(e);
        }

        void CreateStartupShortcut(string path)
        {
            var shell = new WshShellClass();
            var shortcut = (IWshShortcut)shell.CreateShortcut(path);

            var targetPath = Process.GetCurrentProcess().MainModule.FileName;
            var workingDirectory = new FileInfo(targetPath).Directory.FullName;

            shortcut.WorkingDirectory = workingDirectory;
            shortcut.TargetPath = targetPath;
            shortcut.IconLocation = Path.Combine(workingDirectory, @"Assets\keyboard.ico");
            shortcut.Description = "Shortcut to FasType App";

            shortcut.Save();
        }

        void RemoveStartupShortcut(string path)
        {
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);
        }

        public void UpdateStartupShortcut(bool create)
        {
            var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            var shortcutLinkFilePath = Path.Combine(startupFolderPath, FasType.Properties.Resources.AppName + ".lnk");

            if (create)
                CreateStartupShortcut(shortcutLinkFilePath);
            else
                RemoveStartupShortcut(shortcutLinkFilePath);
        }
    }
}
