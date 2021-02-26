using FasType.Utils;
using FasType.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// Interaction logic for PopupWindow.xaml
    /// </summary>
    public partial class PopupWindow : Window
    {
        public static bool IsOpen { get; private set; }

        readonly PopupViewModel _vm;

        static PopupWindow() => IsOpen = false;
        public PopupWindow(PopupViewModel vm)
        {
            InitializeComponent();
            Owner = App.Current.MainWindow;
            DataContext = _vm = vm;

            KeyDown += PopupWindow_KeyDown;

            IsOpen = true;
            Closed += delegate { IsOpen = false; };

            Loaded += delegate
            {
                new System.Media.SoundPlayer(@"Assets\sound.wav").Play();
                this.FlashWindowUntillFocus();
            };
        }

        private void PopupWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        public void SearchForWord(string currentWord) => _vm.SearchForWord(currentWord);
    }
}
