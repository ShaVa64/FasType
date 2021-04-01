using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using FasType.Models;

namespace FasType.ViewModels
{
    public class TaskbarIconViewModel : ObservableObject
    {
        public ICommand ExitApplicationCommand { get; }
        public ICommand ShowAppCommand { get; }

        public TaskbarIconViewModel()
        {
            ExitApplicationCommand = new Command(ExitApplication);
            ShowAppCommand = new Command(ShowApp);
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
