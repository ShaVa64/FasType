using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using FasType.ViewModels;

namespace FasType.Windows
{
    /// <summary>
    /// Interaction logic for OneLettersWindow.xaml
    /// </summary>
    public partial class OneLettersWindow : Window
    {
        public static bool IsOpen { get; private set; }

        static OneLettersWindow() => IsOpen = false;
        public OneLettersWindow(OneLettersViewModel vm)
        {
            InitializeComponent();
            //Owner = App.Current.MainWindow.IsLoaded ? App.Current.MainWindow : null;
            Owner = App.Current.MainWindow;

            IsOpen = true;
            Closed += delegate { IsOpen = false; };

            DataContext = vm;
        }
    }
}
