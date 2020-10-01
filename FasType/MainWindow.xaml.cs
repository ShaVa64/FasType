using FasType.LLKeyboardListener;
using FasType.Utils;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using WindowsInput;
using Serilog;
using FasType.Services;
using Microsoft.Extensions.Options;
using FasType.Abbreviations;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;

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
            var tw = App.Current.ServiceProvider.GetRequiredService<ToolWindow>();
            var p = new Pages.SimpleAbbreviationPage();

            tw.Content = p;

            tw.ShowDialog();
        }
    }
}