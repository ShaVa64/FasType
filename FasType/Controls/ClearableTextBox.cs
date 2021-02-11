using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FasType.Controls
{
    /// <summary>
    /// Follow steps 1a or 1b and then 2 to use this custom control in a XAML file.
    ///
    /// Step 1a) Using this custom control in a XAML file that exists in the current project.
    /// Add this XmlNamespace attribute to the root element of the markup file where it is 
    /// to be used:
    ///
    ///     xmlns:MyNamespace="clr-namespace:FasType.Controls"
    ///
    ///
    /// Step 1b) Using this custom control in a XAML file that exists in a different project.
    /// Add this XmlNamespace attribute to the root element of the markup file where it is 
    /// to be used:
    ///
    ///     xmlns:MyNamespace="clr-namespace:FasType.Controls;assembly=FasType.Controls"
    ///
    /// You will also need to add a project reference from the project where the XAML file lives
    /// to this project and Rebuild to avoid compilation errors:
    ///
    ///     Right click on the target project in the Solution Explorer and
    ///     "Add Reference"->"Projects"->[Browse to and select this project]
    ///
    ///
    /// Step 2)
    /// Go ahead and use your control in the XAML file.
    ///
    ///     <MyNamespace:ClearableTextBox/>
    ///
    /// </summary>
    public class ClearableTextBox : TextBox
    {
        static ClearableTextBox() => DefaultStyleKeyProperty.OverrideMetadata(typeof(ClearableTextBox), new FrameworkPropertyMetadata(typeof(ClearableTextBox)));

        //protected virtual static Type GetCurrentType() => typeof(ClearableTextBox);
        protected static ResourceDictionary DefaultRes => (App.Current.FindResource(typeof(ClearableTextBox)) as Style).Resources;
        protected static T GetResource<T>([CallerMemberName] string name = "") where T : class => DefaultRes[name] as T;
        public static Brush DefaultBrush => GetResource<Brush>();
        //static ClearableTextBox() => _defaultRes = new ResourceDictionary { Source = new(Path.Combine(Directory.GetCurrentDirectory(), @"Dictionaries\TextBoxDictionary.xaml")/*, UriKind.Relative*/) };

        private static readonly DependencyPropertyKey HasTextPropertyKey = DependencyProperty.RegisterReadOnly(nameof(HasText), typeof(bool), typeof(ClearableTextBox), new());
        public static readonly DependencyProperty HasTextProperty = HasTextPropertyKey.DependencyProperty;
        public static readonly DependencyProperty IsClearableProperty = DependencyProperty.Register(nameof(IsClearable), typeof(bool), typeof(ClearableTextBox), new(false));
        public bool IsClearable { get => (bool)GetValue(IsClearableProperty); set => SetValue(IsClearableProperty, value); }
        public bool HasText { get => (bool)GetValue(HasTextProperty); protected set => SetValue(HasTextPropertyKey, value); }
        public ClearableTextBox()
        {
            //Resources.MergedDictionaries.Add(DefaultRes);
            //Template = FindResource("MyTextBoxControlTemplate") as ControlTemplate;

            TextChanged += ClearableTextBox_TextChanged;
            //var btn = GetTemplateChild("ClearButton");
            //btn.Click += ClearTextEvent;

            Loaded += ClearableTextBox_Loaded;

            BorderBrush = DefaultBrush;
        }

        //public override void OnApplyTemplate()
        //{
        //    base.OnApplyTemplate();
        //    BorderBrush = DefaultBrush;
        //}

        private void ClearableTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            var btn = GetTemplateChild("ClearButton") as Button;
            btn.Click += ClearTextEvent;
        }

        private void ClearTextEvent(object sender, RoutedEventArgs e) => Text = string.Empty;
        private void ClearableTextBox_TextChanged(object sender, TextChangedEventArgs e) => HasText = !string.IsNullOrEmpty(Text);
    }
}
