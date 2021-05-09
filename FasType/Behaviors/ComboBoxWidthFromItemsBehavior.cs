using FasType.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace FasType.Behaviors
{
    public class ComboBoxWidthFromItemsBehavior
    {
        public static readonly DependencyProperty ComboBoxWidthFromItemsProperty = DependencyProperty.RegisterAttached("ComboBoxWidthFromItems",
                                                                                                                       typeof(bool),
                                                                                                                       typeof(ComboBoxWidthFromItemsBehavior),
                                                                                                                       new UIPropertyMetadata(false, OnComboBoxWidthFromItemsPropertyChanged));
        
        public static bool GetComboBoxWidthFromItems(DependencyObject obj) => (bool)obj.GetValue(ComboBoxWidthFromItemsProperty);
        public static void SetComboBoxWidthFromItems(DependencyObject obj, bool value) => obj.SetValue(ComboBoxWidthFromItemsProperty, value);
        
        private static void OnComboBoxWidthFromItemsPropertyChanged(DependencyObject dpo, DependencyPropertyChangedEventArgs e)
        {
            if (dpo is ComboBox comboBox)
            {
                if ((bool)e.NewValue == true)
                {
                    comboBox.Loaded += OnComboBoxLoaded;
                    //comboBox.SelectionChanged += ComboBox_SelectionChanged;
                    comboBox.IsVisibleChanged += ComboBox_IsVisibleChanged;
                }
                else
                {

                    comboBox.Loaded -= OnComboBoxLoaded;
                    //comboBox.SelectionChanged -= ComboBox_SelectionChanged;
                    comboBox.IsVisibleChanged -= ComboBox_IsVisibleChanged;
                }
            }
        }


        static void SetComboBoxWidthFromItems(ComboBox cb)
        {
            Action action = cb.SetWidthFromItems;
            cb.Dispatcher.BeginInvoke(DispatcherPriority.Render, action);
        }

        private static void ComboBox_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true)
                SetComboBoxWidthFromItems(sender as ComboBox ?? throw new NullReferenceException());
        }

        private static void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => SetComboBoxWidthFromItems(sender as ComboBox ?? throw new NullReferenceException());
        private static void OnComboBoxLoaded(object? sender, RoutedEventArgs e) => SetComboBoxWidthFromItems(sender as ComboBox ?? throw new NullReferenceException());
    }
}
