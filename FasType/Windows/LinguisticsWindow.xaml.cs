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
    /// Interaction logic for LinguisticsWindow.xaml
    /// </summary>
    public partial class LinguisticsWindow : Window
    {
        readonly LinguisticsViewModel _vm;

        public LinguisticsWindow(LinguisticsViewModel vm)
        {
            InitializeComponent();

            KeyDown += LinguisticsWindow_KeyDown;
            DataContext = _vm = vm;
        }

        private void LinguisticsWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
