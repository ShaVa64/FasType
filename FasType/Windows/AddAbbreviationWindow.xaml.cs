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
    /// Interaction logic for ToolWindow.xaml
    /// </summary>
    public partial class AddAbbreviationWindow : Window
    {
        public static bool IsOpen { get; private set; }

        static AddAbbreviationWindow() => IsOpen = false;
        public AddAbbreviationWindow()
        {
            InitializeComponent();
            Owner = App.Current.MainWindow;
            IsOpen = true;

            Closed += delegate { IsOpen = false; };
        }

        private void AddAbbreviationWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }
    }
}
