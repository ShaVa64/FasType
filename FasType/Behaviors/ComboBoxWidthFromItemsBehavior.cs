﻿using FasType.Utils;
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
                    comboBox.Loaded += OnComboBoxLoaded;
                else
                    comboBox.Loaded -= OnComboBoxLoaded;
            }
        }

        private static void OnComboBoxLoaded(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox ?? throw new NullReferenceException();
            Action action = comboBox.SetWidthFromItems;
            comboBox.Dispatcher.BeginInvoke(DispatcherPriority.Render, action);
        }
    }
}
