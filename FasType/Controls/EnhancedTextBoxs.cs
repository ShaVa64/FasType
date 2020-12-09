using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace FasType.Controls
{
    public class ClearableTextBox : TextBox
    {
        protected readonly static ResourceDictionary _defaultRes;
        protected static T GetResource<T>([CallerMemberName] string name = "") where T : class => _defaultRes[name] as T;
        public static Brush DefaultBrush => GetResource<Brush>();
        static ClearableTextBox() => _defaultRes = new ResourceDictionary { Source = new(@"..\Dictionaries\TextBoxDictionary.xaml", UriKind.Relative) };


        private static readonly DependencyPropertyKey HasTextPropertyKey = DependencyProperty.RegisterReadOnly(nameof(HasText), typeof(bool), typeof(ClearableTextBox), new());
        public static readonly DependencyProperty HasTextProperty = HasTextPropertyKey.DependencyProperty;
        public static readonly DependencyProperty IsClearableProperty = DependencyProperty.Register(nameof(IsClearable), typeof(bool), typeof(ClearableTextBox), new(false));
        public bool IsClearable { get => (bool)GetValue(IsClearableProperty); set => SetValue(IsClearableProperty, value); }
        public bool HasText { get => (bool)GetValue(HasTextProperty); protected set => SetValue(HasTextPropertyKey, value); }
        public ClearableTextBox()
        {
            Resources.MergedDictionaries.Add(_defaultRes);
            Template = FindResource("MyTextBoxControlTemplate") as ControlTemplate;

            TextChanged += ClearableTextBox_TextChanged;
            //var btn = GetTemplateChild("ClearButton");
            //btn.Click += ClearTextEvent;

            Loaded += ClearableTextBox_Loaded;

            BorderBrush = DefaultBrush;
        }

        private void ClearableTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            var btn = GetTemplateChild("ClearButton") as Button;
            btn.Click += ClearTextEvent;
        }

        private void ClearTextEvent(object sender, RoutedEventArgs e) => Text = string.Empty;
        private void ClearableTextBox_TextChanged(object sender, TextChangedEventArgs e) => HasText = !string.IsNullOrEmpty(Text);
    }

    public class BorderBrushTextBox : ClearableTextBox
    {
        #region GroupName
        readonly static Dictionary<string, List<BorderBrushTextBox>> Groups = new();

        public static readonly DependencyProperty GroupNameProperty = DependencyProperty.RegisterAttached("GroupName", typeof(string), typeof(BorderBrushTextBox), new(GroupNameChange));
        public static string GetGroupName(DependencyObject obj) => (string)obj.GetValue(GroupNameProperty);
        public static void SetGroupName(DependencyObject obj, string value) => obj.SetValue(GroupNameProperty, value);

        public static void CheckGroup(string groupName)
        {
            var bbtbs = Groups[groupName].OrderBy(bbtb => bbtb.Text).ToArray();

            int lastSeen = 0;
            for (int i = 0; i < bbtbs.Length; i++)
            {
                if (lastSeen == i)
                    bbtbs[lastSeen].BorderBrush = WarningBrush;

                if (i + 1 < bbtbs.Length && bbtbs[i].Text == bbtbs[i + 1].Text)
                    bbtbs[i + 1].BorderBrush = WarningBrush;
                else
                {
                    if (lastSeen == i)
                        bbtbs[lastSeen].BorderBrush = DefaultBrush;
                    lastSeen = i + 1;
                }
            }
        }

        public static void GroupNameChange(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            var bbtb = obj as BorderBrushTextBox;
            string newGroup = e.NewValue as string;
            string oldGroup = e.OldValue as string;

            if (!string.IsNullOrEmpty(newGroup))
            {
                if (!Groups.ContainsKey(newGroup))
                    Groups.Add(newGroup, new());
                Groups[newGroup].Add(bbtb);
                Groups[newGroup].ForEach(tb => tb.UpdateBorderColor());

                bbtb.Unloaded += (sender, e) => 
                {
                    SetGroupName(sender as DependencyObject, string.Empty);
                };
            }

            if (!string.IsNullOrEmpty(oldGroup))
            {
                if (Groups.ContainsKey(oldGroup))
                    Groups[oldGroup].Remove(bbtb);
                if (Groups[oldGroup].Count == 0)
                    Groups.Remove(oldGroup);
            }
        }
        #endregion
        public static Brush ErrorBrush => GetResource<Brush>();
        public static Brush WarningBrush => GetResource<Brush>();


        public static readonly DependencyProperty ForceBorderBrushProperty = DependencyProperty.Register(nameof(ForcedBorderBrush), typeof(Brush), typeof(BorderBrushTextBox));
        public Brush ForcedBorderBrush { get => (Brush)GetValue(ForceBorderBrushProperty); set => SetValue(ForceBorderBrushProperty, value); }
        
        public BorderBrushTextBox()
        {
            //BorderBrush = string.IsNullOrEmpty(Text) ? ErrorBrush : DefaultBrush;
            UpdateBorderColor();

            TextChanged += BorderBrushTextBox_TextChanged;

            //var style = FindResource("BorderBrushTextBox") as Style;
            //Style = style;
        }

        private void BorderBrushTextBox_TextChanged(object sender, TextChangedEventArgs e) => UpdateBorderColor();
        public void UpdateBorderColor()
        {
            BorderBrush = DefaultBrush;
            string gn = GetGroupName(this);
            
            if (!string.IsNullOrEmpty(gn))
                CheckGroup(gn);
            if (string.IsNullOrEmpty(Text))
                BorderBrush = ErrorBrush;

            BorderBrush = ForcedBorderBrush ?? BorderBrush;
        }
    }
}
