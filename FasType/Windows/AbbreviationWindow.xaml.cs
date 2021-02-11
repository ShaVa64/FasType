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
    public partial class AbbreviationWindow : Window
    {
        public static bool IsOpen { get; private set; }

        static AbbreviationWindow() => IsOpen = false;
        public AbbreviationWindow()
        {
            InitializeComponent();
            Owner = App.Current.MainWindow;
            IsOpen = true;

            Closed += delegate { IsOpen = false; };
        }

        protected override void OnContentChanged(object oldContent, object newContent)
        {
            base.OnContentChanged(oldContent, newContent);
            if (newContent is Page p)
            {
                //p.DataContextChanged += delegate { DataContext = p.DataContext; };
                DataContext = p.DataContext;
            }
        }

        //private void AbbreviationWindow_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Escape)
        //        Close();
        //}
    }
}
