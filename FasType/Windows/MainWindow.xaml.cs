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

            _vm.Load();

            Closing += (s, e) => _vm.Close();
            SizeChanged += MainWindow_SizeChanged;
            ShowActivated = false;
        }

        void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e) => UpdatePos();

        void UpdatePos() => UpdatePos(new((int)Left, (int)Top));
        void UpdatePos(System.Drawing.Point p)
        {
            var wa = Caret.GetWorkingArea(p);

            if (Left < wa.Left)
                Left = wa.Left;
            if (Left + Width > wa.Right)
                Left = wa.Right - Width;

            if (Top < wa.Top)
                Top = wa.Top;
            if (Top + Height > wa.Bottom)
                Top = wa.Bottom - Height;
        }

        public void ShowAt(System.Drawing.Point p)
        {
            Left = p.X;
            Top = p.Y;

            UpdatePos(p);
            Show();
        }
    }
}