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
        public static bool IsOpen { get; private set; }

        static LinguisticsWindow() => IsOpen = false;
        public LinguisticsWindow(LinguisticsViewModel vm)
        {
            InitializeComponent();
            IsOpen = true;

            Closed += delegate { IsOpen = false; };

            //KeyDown += LinguisticsWindow_KeyDown;
            DataContext = vm;
        }

        private void LinguisticsWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
