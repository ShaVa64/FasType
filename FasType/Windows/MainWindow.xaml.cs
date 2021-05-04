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

            //var area = SystemParameters.WorkArea;
            //Left = area.Right - Width;
            //Top = area.Bottom - Height;

            _vm.Load();
            //Loaded += _vm.Load;

            Closing += (s, e) => _vm.Close();
            ShowActivated = false;
            //StateChanged += delegate
            //{
            //    if (WindowState == WindowState.Minimized)
            //    {
            //        WindowState = WindowState.Normal;
            //        Hide();
            //    }
            //};
        }

        public void ShowAt(System.Drawing.Point p, System.Drawing.Rectangle wa)
        {
            Left = p.X;
            Top = p.Y;

            if (Left < wa.Left)
                Left = wa.Left;
            if (Left + Width > wa.Right)
                Left = wa.Right - Width;

            if (Top < wa.Top)
                Top = wa.Top;
            if (Top + Height > wa.Bottom)
                Top = wa.Bottom - Height;
            Show();
            //Left = area.Right - Width;
            //Top = area.Bottom - Height;
        }
    }
}