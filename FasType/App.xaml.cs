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
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;
using System.Globalization;
using Microsoft.EntityFrameworkCore;
using System.Windows.Data;
using FasType.Models.Linguistics;

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

        static App() => FrameworkElement.LanguageProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.Name)));
        public App()
        {
            Configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            ServiceProvider = serviceCollection.BuildServiceProvider();

            using var _abbrevContext = ServiceProvider.GetRequiredService<IAbbreviationStorage>() as DbContext;
            _abbrevContext?.Database.Migrate();
            using var _lingContext = ServiceProvider.GetRequiredService<ILinguisticsStorage>() as DbContext;
            _lingContext?.Database.Migrate();
            //_context.Database.EnsureCreated();
            //var x = System.Text.Json.JsonSerializer.Serialize(_lingContext as ILinguisticsStorage/*, new() { WriteIndented = true }*/);

            //T();
            //F();
        }

        private void ConfigureServices(ServiceCollection services)
        {
            Log.Logger = new LoggerConfiguration()
                //.ReadFrom.Configuration(Configuration)
                .MinimumLevel.Verbose()
                .WriteTo.Debug()
                .CreateLogger();

            services.AddDbContext<IAbbreviationStorage, EFSqliteAbbreviationContext>(options => options.UseSqlite(Configuration.GetConnectionString("EFAbbreviation")), ServiceLifetime.Transient, ServiceLifetime.Transient);
            services.AddDbContext<ILinguisticsStorage, EFSqliteLinguisticsContext>(options => options.UseSqlite(Configuration.GetConnectionString("EFLinguistics")), ServiceLifetime.Transient, ServiceLifetime.Transient);
            //services.AddTransient<IDataStorage, EFSqliteStorage>();

            services.AddSingleton(Configuration);

            services.AddSingleton<MainWindow>();
            services.AddTransient<MainWindowViewModel>();
            
            services.AddTransient<AddAbbreviationWindow>();
            services.AddTransient<SimpleAbbreviationPage>();
            services.AddTransient<SimpleAbbreviationViewModel>();

            services.AddTransient<SeeAllWindow>();
            services.AddTransient<SeeAllViewModel>();

            services.AddTransient<LinguisticsWindow>();
            services.AddTransient<LinguisticsViewModel>();

            services.AddTransient<AbbreviationMethodsWindow>();
            services.AddTransient<AbbreviationMethodsViewModel>();
        }

        private void OnStartup(object sender, StartupEventArgs args)
        {
            var mw = ServiceProvider.GetService<MainWindow>();
            MainWindow = mw;
            MainWindow.Show();
        }

        //void T()
        //{
        //    string fp = @"D:\Visual Studio Projects\FasType\Docs\abréviations.txt";
        //    string tb = @"D:\Visual Studio Projects\FasType\Docs\table_mère avec toutes les formes.txt";

        //    using var stream1 = new FileStream(fp, FileMode.Open, FileAccess.Read);
        //    using var reader1 = new StreamReader(stream1); 
            
        //    using var stream2 = new FileStream(tb, FileMode.Open, FileAccess.Read);
        //    using var reader2 = new StreamReader(stream2);

        //    using var _context = ServiceProvider.GetRequiredService<IAbbreviationStorage>();


        //    var abbrevs = new List<(string, string)>();
        //    string l;
        //    while ((l = reader1.ReadLine()) != null)
        //    {
        //        var sp = l.Replace("\"", string.Empty).Split(';');

        //        string sf = sp[0];
        //        string ff = sp[1];

        //        abbrevs.Add((sf, ff));
        //        //_context.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff));
        //    }

        //    var simples = new List<Models.Abbreviations.BaseAbbreviation>();
        //    l = reader2.ReadLine();
        //    while ((l = reader2.ReadLine()) != null)
        //    {
        //        var sp = l.Replace("\"", string.Empty).Split(';');

        //        string ff = sp[1];
        //        string gf = sp[2];
        //        string gpf = sp[3];
        //        string pf = sp[4];

        //        string sf;
        //        var sfs = abbrevs.Where(t => t.Item2 == ff).ToArray();

        //        if (sfs.Length == 1)
        //            sf = sfs.Single().Item1;
        //        else
        //        {
        //            continue;
        //        }

        //        simples.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff, 0, gf, pf, gpf));
        //        //_context.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff));
        //    }
        //    _context.AddRange(simples);
        //}

        //void F()
        //{
        //    string fp = @"D:\Visual Studio Projects\FasType\Docs\méthode abréviation.txt";

        //    using var stream = new FileStream(fp, FileMode.Open, FileAccess.Read);
        //    using var reader = new StreamReader(stream);

        //    using var _context = ServiceProvider.GetRequiredService<ILinguisticsStorage>();


        //    var methods = new List<SyllableAbbreviation>();
        //    string l = reader.ReadLine();
        //    while ((l = reader.ReadLine()) != null)
        //    {
        //        var sp = l.Replace("\"", string.Empty).Split(';');

        //        string ff = sp[0];
        //        string sf = sp[1];
        //        SyllablePosition p = SyllablePosition.None;
        //        if (sp[2] == "1")
        //            p |= SyllablePosition.Before;
        //        if (sp[3] == "1")
        //            p |= SyllablePosition.In;
        //        if (sp[4] == "1")
        //            p |= SyllablePosition.After;

        //        methods.Add(new SyllableAbbreviation(Guid.NewGuid(), sf, ff, p));
        //        //_context.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff));
        //    }
        //    _context.AbbreviationMethods = methods;

        //    var x = System.Text.Json.JsonSerializer.Serialize(_context);
        //    var xx = System.Text.Json.JsonSerializer.Deserialize<LinguisticsDTO>(x);
        //    //var x = System.Text.Json.JsonSerializer.Serialize(methods);
        //}

        protected override void OnExit(ExitEventArgs e)
        {
            FasType.Properties.Settings.Default.Save();
            Log.Information("Default Settings saved!");
            base.OnExit(e);
        }

        static void CreateStartupShortcut(string path)
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

        static void RemoveStartupShortcut(string path)
        {
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);
        }

        public static void UpdateStartupShortcut(bool shouldCreate)
        {
            var startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            var shortcutLinkFilePath = Path.Combine(startupFolderPath, FasType.Properties.Resources.AppName + ".lnk");

            if (shouldCreate)
                CreateStartupShortcut(shortcutLinkFilePath);
            else
                RemoveStartupShortcut(shortcutLinkFilePath);
        }
    }
}
