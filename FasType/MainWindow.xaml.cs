using System.Windows;
using FasType.Services;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace FasType
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly IKeyboardListenerHandler _listenerHandler;

        public MainWindow(IKeyboardListenerHandler listenerHandler)
        {
            InitializeComponent();
            _listenerHandler = listenerHandler;

            var area = System.Windows.SystemParameters.WorkArea;
            this.Left = area.Right - this.Width;
            this.Top = area.Bottom - this.Height;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e) => _listenerHandler.Load(CurrentWordCallback);
        void CurrentWordCallback(string currentWord) => label.Text = currentWord;
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) => _listenerHandler.Close();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not FrameworkElement fe)
                return;

            if (fe.Tag is not Type t)
                return;

            var tw = App.Current.ServiceProvider.GetRequiredService<ToolWindow>();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as Page;// Activator.CreateInstance(t) as Page;//new Pages.SimpleAbbreviationPage();

            tw.Content = p;

            _listenerHandler.Pause();
            tw.ShowDialog();
            _listenerHandler.Continue();
        }
    }
}