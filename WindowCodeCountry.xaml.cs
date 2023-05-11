using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ExcelToTxt
{
    /// <summary>
    /// Interaction logic for WindowCodeCountry.xaml
    /// </summary>
    public partial class WindowCodeCountry : Window
    {
        public CountryCode CountryCode { get; set; }

        private List<CountryCode> listCodes = new List<CountryCode>();
        public WindowCodeCountry()
        {
            InitializeComponent();
        }

        private void Load_Codes_Country(object sender, EventArgs e)
        {
            listCodes.Add(new CountryCode() { Code = 1, Country = "AUSTRIA" });
            listCodes.Add(new CountryCode() { Code = 2, Country = "CZECH_REPUBLIC" });
            listCodes.Add(new CountryCode() { Code = 3, Country = "SLOVAKIA" });
            listCodes.Add(new CountryCode() { Code = 4, Country = "HUNGARY" });
            listCodes.Add(new CountryCode() { Code = 5, Country = "POLAND" });
            listCodes.Add(new CountryCode() { Code = 6, Country = "ROMANIA" });
            listCodes.Add(new CountryCode() { Code = 7, Country = "GREECE" });
            listCodes.Add(new CountryCode() { Code = 8, Country = "BULGARIA" });
            listCodes.Add(new CountryCode() { Code = 10, Country = "TURKEY" });
            listCodes.Add(new CountryCode() { Code = 11, Country = "UKRAINE" });

            foreach (var item in listCodes)
            {
                comboBoxCodeCountry.Items.Add("Code: " + item.Code.ToString() + ", Country: " + item.Country);
            }

        }

        private void ComboBoxCodeCountry_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = comboBoxCodeCountry.SelectedIndex;
            CountryCode = listCodes.ElementAt(index);
        }

        private void Button_Click_Submit(object sender, RoutedEventArgs e)
        {
            if (CountryCode != null) this.Hide();
            else MessageBox.Show("Please select country code and select button Submit !!!");
        }
    }
}
