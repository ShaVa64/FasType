using FasType.ViewModels;
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

namespace FasType.Windows
{
    /// <summary>
    /// Interaction logic for SyllableAbbreviationWindow.xaml
    /// </summary>
    public partial class AbbreviationMethodsWindow : Window
    {
        public static bool IsOpen { get; private set; }

        static AbbreviationMethodsWindow() => IsOpen = false;
        public AbbreviationMethodsWindow(AbbreviationMethodsViewModel vm)
        {
            InitializeComponent();
            IsOpen = true;
            Owner = App.Current.MainWindow.IsLoaded ? App.Current.MainWindow : null;

            Closed += delegate { IsOpen = false; };
            //KeyDown += LinguisticsWindow_KeyDown;
            DataContext /*= _vm */= vm;
        }

        private void LinguisticsWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
