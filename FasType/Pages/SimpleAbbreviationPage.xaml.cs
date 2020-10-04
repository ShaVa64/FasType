using FasType.Models;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FasType.Pages
{
    /// <summary>
    /// Interaction logic for SimpleAbbreviationPage.xaml
    /// </summary>
    public partial class SimpleAbbreviationPage : Page
    {
        readonly IDataStorage _storage;
        SimpleAbbreviation _currentAbbrev;

        public SimpleAbbreviationPage(IDataStorage storage)
        {
            InitializeComponent();

            _storage = storage;
            _currentAbbrev = null;
            sfTB.Focus();
            //Init();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            if (_currentAbbrev == null || string.IsNullOrEmpty(_currentAbbrev.ShortForm) || string.IsNullOrEmpty(_currentAbbrev.FullForm))
            {
                MessageBox.Show("You can't create an empty abbreviation", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            bool b = await _storage.AddAsync(_currentAbbrev);
            if (!b)
            {
                MessageBox.Show($"An error has occured while trying to create the abbreviation ({_currentAbbrev.ElementaryRepresentation}).", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            (Parent as Window).Close();
        }
        
        private void TB_TextChanged(object sender, TextChangedEventArgs e)
        {
            PreviewText.Text = "";

            string sf = sfTB.Text.Trim();
            string ff = ffTB.Text.Trim();
            if (string.IsNullOrEmpty(sf) && string.IsNullOrEmpty(ff))
                return;

            _currentAbbrev = new SimpleAbbreviation(sf, ff);

            PreviewText.Text = _currentAbbrev.ComplexRepresentation;
        }

        //void Init()
        //{
        //    Grid g = NewGrid();

        //    var ctors = typeof(Models.SimpleAbbreviation).GetConstructors();
        //    var @params = ctors[0].GetParameters();

        //    foreach (var param in @params)
        //    {
        //        var rd = g.RowDefinitions;
        //        rd.Insert(rd.Count - 1, new RowDefinition() { Height = new GridLength(1, GridUnitType.Star) });
        //        g.Children.Add(NewTextBox(rd.Count - 2));
        //        g.Children.Add(NewLabel(rd.Count - 2, param));
        //    }

        //    g.Children.Add(NewButton(g.RowDefinitions.Count - 1));

        //    Content = g;
        //}

        //Grid NewGrid()
        //{
        //    Grid g = new();

        //    var cd = g.ColumnDefinitions;
        //    cd.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
        //    cd.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });

        //    var rd = g.RowDefinitions;
        //    rd.Add(new RowDefinition() { Height = new GridLength(1, GridUnitType.Star) });

        //    return g;
        //}

        //Button NewButton(int row)
        //{
        //    Button b = new();
        //    b.Content = "Create";

        //    Grid.SetRow(b, row);
        //    Grid.SetColumnSpan(b, 2);

        //    return b;
        //}

        //TextBox NewTextBox(int row)
        //{
        //    TextBox tb = new();
        //    Grid.SetColumn(tb, 1);
        //    Grid.SetRow(tb, row);

        //    return tb;
        //}

        //Label NewLabel(int row, ParameterInfo pi)
        //{
        //    Label l = new();
        //    Grid.SetColumn(l, 0);
        //    Grid.SetRow(l, row);

        //    l.Content = string.Concat(pi.Name.Select((c, i) => i > 0 ? ((char.IsUpper(c) ? " " : "") + c) : char.ToUpper(c).ToString())) + ":";

        //    return l;
        //}
    }
}
