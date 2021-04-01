using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace FasType.Controls
{
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
            var bbtb = obj as BorderBrushTextBox ?? throw new NullReferenceException();
            string newGroup = e.NewValue as string ?? throw new NullReferenceException();
            string oldGroup = e.OldValue as string ?? throw new NullReferenceException();

            if (!string.IsNullOrEmpty(newGroup))
            {
                if (!Groups.ContainsKey(newGroup))
                    Groups.Add(newGroup, new());
                Groups[newGroup].Add(bbtb);
                Groups[newGroup].ForEach(tb => tb.UpdateBorderColor());

                bbtb.Unloaded += (sender, e) => 
                {
                    SetGroupName(sender as DependencyObject ?? throw new NullReferenceException(), string.Empty);
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

        //protected static new T GetResource<T>([CallerMemberName] string name = "") where T : class => (App.Current.FindResource(typeof(BorderBrushTextBox)) as Style).Resources[name] as T;
        public static Brush ErrorBrush => GetResource<Brush>();
        public static Brush WarningBrush => GetResource<Brush>();


        public static readonly DependencyProperty ForcedBorderBrushProperty = DependencyProperty.Register(nameof(ForcedBorderBrush), typeof(Brush), typeof(BorderBrushTextBox));
        public Brush ForcedBorderBrush { get => (Brush)GetValue(ForcedBorderBrushProperty); set => SetValue(ForcedBorderBrushProperty, value); }

        //protected static new Type CurrentType => typeof(BorderBrushTextBox);
        static BorderBrushTextBox() => DefaultStyleKeyProperty.OverrideMetadata(typeof(BorderBrushTextBox), new FrameworkPropertyMetadata(typeof(BorderBrushTextBox)));

        public BorderBrushTextBox()
        {
            //BorderBrush = string.IsNullOrEmpty(Text) ? ErrorBrush : DefaultBrush;
            UpdateBorderColor();

            TextChanged += BorderBrushTextBox_TextChanged;

            //var style = FindResource("BorderBrushTextBox") as Style;
            //Style = style;
        }

        //public override void OnApplyTemplate()
        //{
        //    base.OnApplyTemplate();
        //    UpdateBorderColor();
        //}

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
