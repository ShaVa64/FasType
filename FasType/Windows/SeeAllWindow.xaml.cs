using FasType.ViewModels;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using FasType.Models.Abbreviations;

namespace FasType.Windows
{
    /// <summary>
    /// Interaction logic for SeeAllWindow.xaml
    /// </summary>
    public partial class SeeAllWindow : Window
    {
        public static bool IsOpen { get; private set; }

        static SeeAllWindow() => IsOpen = false;
        public SeeAllWindow(SeeAllViewModel vm)
        {
            InitializeComponent();
            //Owner = App.Current.MainWindow.IsLoaded ? App.Current.MainWindow : null;
            Owner = App.Current.MainWindow;

            IsOpen = true;
            Closed += delegate { IsOpen = false; };

            DataContext = vm;
        }

        private void SeeAllWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
