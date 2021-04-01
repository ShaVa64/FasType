using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Controls;

namespace FasType.Selectors
{
    public class OrderSelector : DataTemplateSelector
    {
        [NotNull] public DataTemplate? First { get; set; }
        [NotNull] public DataTemplate? Default { get; set; }
        [NotNull] public DataTemplate? Last { get; set; }
        [NotNull] public DataTemplate? Only { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            int altIndex = ItemsControl.GetAlternationIndex(container);

            var ic = ItemsControl.ItemsControlFromItemContainer(container);
            int altCount = ic.AlternationCount;

            if (altCount == 1)
                return Only;

            if (altIndex == 0)
                return First;
            if (altIndex == altCount - 1)
                return Last;
            return Default;
        }
    }
}
