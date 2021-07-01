using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using FasType.Core.Models;
using FasType.Core.Services;
using FasType.Models;
using FasType.Windows;
using Microsoft.Extensions.DependencyInjection;

namespace FasType.ViewModels
{
    public class TaskbarIconViewModel : ObservableObject
    {
        private readonly IRepositoriesManager _repositories;

        public ICommand ExitApplicationCommand { get; }
        public ICommand AddNewCommand { get; }
        public ICommand SeeAllCommand { get; }
        public ICommand OpenLinguisticsCommand { get; }

        public TaskbarIconViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;
            ExitApplicationCommand = new Command(ExitApplication);

            AddNewCommand = new Command<Type>(AddNew, CanAddNew);
            SeeAllCommand = new Command(SeeAll, CanSeeAll);
            OpenLinguisticsCommand = new Command(OpenLinguistics, CanOpenLinguistics);
        }
        private bool CanOpenLinguistics() => !LinguisticsWindow.IsOpen;
        private void OpenLinguistics()
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

        bool CanSeeAll()
        {
            bool c = _repositories.Abbreviations.Count > 0;
            _repositories.Reload();
            return c && !SeeAllWindow.IsOpen;
        }

        void SeeAll()
        {
            var saw = App.Current.ServiceProvider.GetRequiredService<SeeAllWindow>();

            saw.Show();
        }
        void ExitApplication()
        {
            App.Current.Shutdown();
        }
    }
}
