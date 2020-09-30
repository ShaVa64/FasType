﻿using FasType.LLKeyboardListener;
using FasType.Services;
using FasType.Storage;
using IWshRuntimeLibrary;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
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
                .ReadFrom.Configuration(Configuration)
                //.MinimumLevel.Verbose()
                //.WriteTo.Debug()
                .CreateLogger();

            services.AddSingleton(Configuration);
            services.AddSingleton<MainWindow>();
            services.AddSingleton<IDataStorage>(new FileDataStorage(Configuration["DataFilePath"]));
            services.AddTransient<IKeyboardListenerHandler, KeyboardListenerHandler>();
            services.AddTransient<ToolWindow>();
        }

        private void OnStartup(object sender, StartupEventArgs args)
        {
            Task.Run(CheckStartupShortcut);

            MainWindow = ServiceProvider.GetRequiredService<MainWindow>();
            MainWindow.Show();
        }

        void CheckStartupShortcut()
        {
            if (FasType.Properties.Settings.Default.OnStartUp is false)
                return;

            var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            var shell = new WshShellClass();

            var shortcutLinkFilePath = Path.Combine(startupFolderPath, FasType.Properties.Resources.AppName + ".lnk");

            var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLinkFilePath);

            var targetPath = Process.GetCurrentProcess().MainModule.FileName;
            var workingDirectory = new FileInfo(targetPath).Directory.FullName;

            shortcut.WorkingDirectory = workingDirectory;
            shortcut.TargetPath = targetPath;

            shortcut.Save();
        }

    }
}
