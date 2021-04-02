using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using FasType.Models;
using FasType.Windows;
using Microsoft.Extensions.DependencyInjection;

namespace FasType.ViewModels
{
    public class TaskbarIconViewModel : ObservableObject
    {
        public ICommand ExitApplicationCommand { get; }
        public ICommand ShowAppCommand { get; }
        public ICommand AddNewCommand { get; }
        public ICommand SeeAllCommand { get; }
        public ICommand OpenLinguisticsCommand { get; }

        public TaskbarIconViewModel()
        {
            ExitApplicationCommand = new Command(ExitApplication);
            ShowAppCommand = new Command(ShowApp);

            AddNewCommand = new Command<Type>(AddNew, CanAddNew);
            SeeAllCommand = new Command(SeeAll, CanSeeAll);
            OpenLinguisticsCommand = new Command(OpenLinguistics, CanOpenLinguistics);
        }
        bool CanOpenLinguistics() => !LinguisticsWindow.IsOpen;
        void OpenLinguistics()
        {
            var lw = App.Current.ServiceProvider.GetRequiredService<LinguisticsWindow>();

            lw.Show();
        }

        bool CanAddNew(Type? t) => t != null && t.IsSubclassOf(typeof(Page)) && !AbbreviationWindow.IsOpen;
        void AddNew(Type? t)
        {
            _ = t ?? throw new NullReferenceException();
            var aaw = App.Current.ServiceProvider.GetRequiredService<AbbreviationWindow>();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as Page;

            aaw.Content = p;
            aaw.Show();
        }

        bool CanSeeAll() => App.Current.ServiceProvider.GetRequiredService<Services.IAbbreviationStorage>().Count > 0 && !SeeAllWindow.IsOpen;
        void SeeAll()
        {
            var saw = App.Current.ServiceProvider.GetRequiredService<SeeAllWindow>();

            saw.Show();
        }

        void ShowApp()
        {
            App.Current.MainWindow.Show();
        }
        void ExitApplication()
        {
            App.Current.Shutdown();
        }
    }
}
