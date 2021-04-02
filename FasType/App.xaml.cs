using FasType.LLKeyboardListener;
using FasType.Windows;
using FasType.Pages;
using FasType.Services;
using FasType.Storage;
using FasType.ViewModels;
using IWshRuntimeLibrary;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
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
using System.Text;
using Hardcodet.Wpf.TaskbarNotification;

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

        TaskbarIcon? taskbarIcon;

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
            using var _dicContext = ServiceProvider.GetRequiredService<IDictionaryStorage>() as DbContext;
            _dicContext?.Database.Migrate();
            //_context.Database.EnsureCreated();
            //var x = System.Text.Json.JsonSerializer.Serialize(_lingContext as ILinguisticsStorage/*, new() { WriteIndented = true }*/);

            //I();
            //H();
            //G();
            //T();
            //F();

            //var res = ServiceProvider.GetRequiredService<ILinguisticsStorage>().Words("pçé");

            //var bbtb = new FasType.Controls.BorderBrushTextBox();
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
            services.AddDbContext<IDictionaryStorage, EFSqliteDictionaryContext>(options => options.UseSqlite(Configuration.GetConnectionString("EFDictionary")), ServiceLifetime.Transient, ServiceLifetime.Transient);
            //services.AddTransient<IDataStorage, EFSqliteStorage>();

            services.AddSingleton(Configuration);

            services.AddSingleton<MainWindow>();
            services.AddTransient<MainWindowViewModel>();
            
            services.AddTransient<AbbreviationWindow>();
            services.AddTransient<SimpleAbbreviationPage>();
            services.AddTransient<AddSimpleAbbreviationViewModel>();
            services.AddTransient<ModifySimpleAbbreviationViewModel>();

            services.AddTransient<SeeAllWindow>();
            services.AddTransient<SeeAllViewModel>();

            services.AddTransient<LinguisticsWindow>();
            services.AddTransient<LinguisticsViewModel>();

            services.AddTransient<AbbreviationMethodsWindow>();
            services.AddTransient<AbbreviationMethodsViewModel>();

            services.AddTransient<OneLettersWindow>();
            services.AddTransient<OneLettersViewModel>();

            services.AddTransient<PopupWindow>();
            services.AddTransient<PopupViewModel>();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            taskbarIcon = (TaskbarIcon)FindResource("NotifyIcon");
            //taskbarIcon.ShowBalloonTip("TEST", "This is a test.", BalloonIcon.None);

            var mw = ServiceProvider.GetRequiredService<MainWindow>();
            MainWindow = mw;
            //MainWindow.Show();
        }

        #region DB Methods
        //void I()
        //{
        //    using var _context = ServiceProvider.GetRequiredService<IDictionaryStorage>() as EFSqliteDictionaryContext;
        //    var l = new List<Models.Dictionary.BaseDictionaryElement>();
        //    string fp = @"D:\Visual Studio Projects\FasType\Docs\Lexique383.tsv";

        //    var lines = System.IO.File.ReadAllLines(fp).Skip(1).ToArray();
        //    var xs = lines.Select(s => s.Split('\t')).Select(t => new { F = t[0], B = t[2], T = t[3], G = t[4], N = t[5] }).ToList();

        //    var ks = xs.Select(x => x.T).Distinct().ToList();
        //    var rs = xs.Where(x => x.T != "VER" && x.T != "AUX").ToList();

        //    var ts = rs.Select(x => x.T).Distinct().Select(t => new { T = t, N = rs.Count(x => x.T == t) }).OrderByDescending(x => x.N).Select(x => $"{x.T}, ({x.N})").ToList();

        //    var gs = rs.GroupBy(x => x.B + "," + x.T).ToList();
        //    var os = gs.Where(g => g.Count() >= 5).ToList();

        //    //gs = gs.Where(g => g.Count() < 5).ToList();

        //    var keys = gs.Select(g => g.Key.Split(',')[0]).ToList();
        //    var ggs = rs.Where(a => keys.Contains(a.B)).GroupBy(a => a.B).ToList();

        //    var ds = ggs.Where(gg => gg.Select(a => a.T).Distinct().Count() >= 2).ToList();
        //    var ss = ggs.Where(gg => gg.Select(a => a.T).Distinct().Count() < 2).ToList();

        //    var r = ss.Where(g => g.Key.Contains("complot")).ToList();

        //    var solos = ss.Where(g => g.Key.Distinct().Count() == 1).ToList();

        //    var vs = ggs.Where(gg => gg.Any(a => a.T.Contains("PRO"))).ToList();
        //}

        //void H()
        //{
        //    using var _context = ServiceProvider.GetRequiredService<IDictionaryStorage>() as EFSqliteDictionaryContext;

        //    var l = new List<Models.Dictionary.BaseDictionaryElement>();
        //    string fp = @"D:\Visual Studio Projects\FasType\Docs\table_mère avec toutes les formes.txt";

        //    var lines = System.IO.File.ReadAllLines(fp);
        //    for (int index = 1; index < lines.Length; index++)
        //    {
        //        var line = lines[index];
        //        var sp = line.Replace("\"", string.Empty).Split(';');

        //        var ff = sp[1];
        //        var gf = sp[2];
        //        var pf = sp[4];
        //        var gpf = sp[3];

        //        if (l.Any(e => e.FullForm == ff))
        //            continue;
        //        if (ff == "passé")
        //            continue;
        //        if (ff.Contains(' '))
        //        {
        //            var sp2 = ff.Split(' ');
        //            if (sp2.All(s => char.IsUpper(s[0])))
        //                continue;
        //        }

        //        var elem = new Models.Dictionary.SimpleDictionaryElement(ff, gf, pf, gpf);
        //        l.Add(elem);
        //    }

        //    _context.Dictionary.AddRange(l);
        //    _context.SaveChanges();

        //    //_context.Add(new Models.Dictionary.SimpleDictionaryElement("passé", "passée", "passés", "passées"));
        //}

        //void G()
        //{
        //    using var _context = ServiceProvider.GetRequiredService<IAbbreviationStorage>() as EFSqliteAbbreviationContext;
        //    var x1 = _context.Abbreviations.Where(ba => ba.ShortForm.Length == 1).ToArray();
        //    var x2 = _context.Abbreviations.Where(ba => ba.FullForm.Contains(" ")).ToArray();
        //}

        //void T()
        //{
        //    string fp = @"D:\Visual Studio Projects\FasType\Docs\abréviations.txt";
        //    string tb = @"D:\Visual Studio Projects\FasType\Docs\table_mère avec toutes les formes.txt";
        //    string db = @"D:\Visual Studio Projects\FasType\Docs\doublons.txt";

        //    using var stream1 = new FileStream(fp, FileMode.Open, FileAccess.Read);
        //    using var reader1 = new StreamReader(stream1);

        //    using var stream2 = new FileStream(tb, FileMode.Open, FileAccess.Read);
        //    using var reader2 = new StreamReader(stream2);

        //    //using var _context = ServiceProvider.GetRequiredService<IAbbreviationStorage>();

        //    StringBuilder sb = new();

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

        //    //var simples = new List<Models.Abbreviations.BaseAbbreviation>();
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

        //        //if (sfs.Length == 1)
        //        //    sf = sfs.Single().Item1;
        //        //else 
        //        if (sfs.Length > 1)
        //        {
        //            sb.Append(ff)
        //                .Append(": ");
        //            sb.AppendJoin(", ", sfs.Select(t => t.Item1));
        //            //foreach (string ssf in sfs.Select(t => t.Item1))
        //            //    sb.Append(ssf).Append(", ");
        //            sb.AppendLine();
        //            string prev = sb.ToString();
        //            continue;
        //        }
        //        else 
        //        {
        //            continue;
        //        }

        //        //simples.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff, 0, gf, pf, gpf));
        //        //_context.Add(new Models.Abbreviations.SimpleAbbreviation(sf, ff));
        //    }
        //    //_context.AddRange(simples);

        //    System.IO.File.WriteAllText(db, sb.ToString());
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
        #endregion

        protected override void OnExit(ExitEventArgs e)
        {
            FasType.Properties.Settings.Default.Save();
            Log.Information("Default Settings saved!");
            
            taskbarIcon?.Dispose();
            base.OnExit(e);
        }

        static void CreateStartupShortcut(string path)
        {
            var shell = new WshShellClass();
            var shortcut = (IWshShortcut)shell.CreateShortcut(path);

            var targetPath = Process.GetCurrentProcess().MainModule?.FileName ?? throw new NullReferenceException();
            var workingDirectory = new FileInfo(targetPath).Directory?.FullName ?? throw new NullReferenceException();

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
