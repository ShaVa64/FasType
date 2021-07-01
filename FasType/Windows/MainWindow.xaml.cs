using System.Windows;
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
            IsVisibleChanged += MainWindow_IsVisibleChanged;
            ShowActivated = false;

            //Causes window to be loaded; won't have to load first time appearing causing slowdown
            Show();
            Hide();
        }

        private void MainWindow_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue is bool b && b)
            {
                BeUnderCaret();
            }
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            UpdatePos();
            base.OnRenderSizeChanged(sizeInfo);
        }
        void UpdatePos() 
        {
            var wa = Caret.GetWorkingArea(new((int)Left, (int)Top));

            if (Left < wa.Left)
                Left = wa.Left;
            if (Left + Width > wa.Right)
                Left = wa.Right - Width;

            if (Top < wa.Top)
                Top = wa.Top;
            if (Top + Height > wa.Bottom)
                Top = wa.Bottom - Height;
        }

        void BeAt(System.Drawing.Point p)
        {
            Left = p.X;
            Top = p.Y;
        }

        public void BeUnderCaret() => BeAt(Caret.GetCaretPos());
    }
}