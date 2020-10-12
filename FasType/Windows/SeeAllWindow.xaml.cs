using FasType.ViewModels;
using System;
using System.Collections.Generic;
using System.Text;
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
    /// Interaction logic for SeeAllWindow.xaml
    /// </summary>
    public partial class SeeAllWindow : Window
    {
        readonly SeeAllViewModel _vm;

        public SeeAllWindow(SeeAllViewModel vm)
        {
            InitializeComponent();
            Owner = App.Current.MainWindow;

            DataContext = _vm = vm;
        }
    }
}
