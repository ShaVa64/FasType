using System.Windows;
using FasType.Services;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using System;
using FasType.Utils;
using FasType.ViewModels;
using System.Windows.Input;
using System.Windows.Ink;
using System.Windows.Interop;
using System.Runtime.InteropServices;


namespace FasType.Windows
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {        
        readonly MainWindowViewModel _vm;

        public MainWindow(MainWindowViewModel vm)
        {
            InitializeComponent();

            DataContext = _vm = vm;

            var area = SystemParameters.WorkArea;
            Left = area.Right - Width;
            Top = area.Bottom - Height;


            _vm.Load(this, new RoutedEventArgs());
            //Loaded += _vm.Load;
            Closing += _vm.Close;

            StateChanged += delegate
            {
                if (WindowState == WindowState.Minimized)
                {
                    WindowState = WindowState.Normal;
                    Hide();
                }
            };
        }
    }
}