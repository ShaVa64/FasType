using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace FasType.Controls
{
    public class CharacterCasingTextBlock : TextBlock
    {
        public static readonly DependencyProperty CharacterCasingProperty = DependencyProperty.Register(nameof(CharacterCasing),
                                                                                                        typeof(CharacterCasing),
                                                                                                        typeof(CharacterCasingTextBlock),
                                                                                                        new PropertyMetadata(CharacterCasing.Normal, CharacterCasingChanged));

        public CharacterCasing CharacterCasing
        {
            get => (CharacterCasing)GetValue(CharacterCasingProperty);
            set => SetValue(CharacterCasingProperty, value);
        }

        static void CharacterCasingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var tb = d as TextBlock;
            tb.Text = (CharacterCasing)e.NewValue switch
            {
                CharacterCasing.Normal => tb.Text,
                CharacterCasing.Upper => tb.Text.ToUpper(),
                CharacterCasing.Lower => tb.Text.ToLower(),
                _ => throw new NotImplementedException()
            };
        }
    }
}
