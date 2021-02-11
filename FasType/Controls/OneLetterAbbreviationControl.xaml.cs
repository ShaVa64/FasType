using FasType.Models;
using FasType.Models.Abbreviations;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FasType.Controls
{
    /// <summary>
    /// Interaction logic for OneLetterAbbreviationControl.xaml
    /// </summary>
    public partial class OneLetterAbbreviationControl : UserControl
    {
        public readonly static DependencyProperty ModifyCommandProperty = DependencyProperty.Register(nameof(ModifyCommand),
                                                                                                      typeof(Command<BaseAbbreviation>),
                                                                                                      typeof(OneLetterAbbreviationControl));
        
        public Command<BaseAbbreviation> ModifyCommand
        {
            get => (Command<BaseAbbreviation>)GetValue(ModifyCommandProperty);
            set => SetValue(ModifyCommandProperty, value);
        }
        
        public OneLetterAbbreviationControl()
        {
            InitializeComponent();
        }
    }
}
