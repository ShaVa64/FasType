using System.Windows;
using FasType.Services;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using System;
using FasType.ViewModels;
using System.Windows.Input;
using System.Windows.Ink;

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

            var area = System.Windows.SystemParameters.WorkArea;
            this.Left = area.Right - this.Width;
            this.Top = area.Bottom - this.Height;

            this.Loaded += _vm.Load;
            this.Closing += _vm.Close;
        }
    }
}