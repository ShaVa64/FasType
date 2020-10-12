using System.Windows;
using FasType.Services;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using System;
using FasType.ViewModels;

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

            this.CommandBindings.Add(new(_vm.AddNewCommand, _vm.AddNew, _vm.CanAddNew));
            this.CommandBindings.Add(new(_vm.SeeAllCommand, _vm.SeeAll, _vm.CanSeeAll));

            var area = System.Windows.SystemParameters.WorkArea;
            this.Left = area.Right - this.Width;
            this.Top = area.Bottom - this.Height;

            this.Loaded += _vm.Load;
            this.Closing += _vm.Close;
        }
    }
}