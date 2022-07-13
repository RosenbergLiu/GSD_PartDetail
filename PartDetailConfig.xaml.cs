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
using System.Windows.Shapes;

namespace PartDetail
{
    /// <summary>
    /// Interaction logic for PartDetailConfig.xaml
    /// </summary>
    public partial class PartDetailConfig : Window
    {
        public PartDetailConfig()
        {
            InitializeComponent();
            SAGusername.Text = Properties.Settings.Default.USERNAME;
            SAGpassword.Text = Properties.Settings.Default.PASSWORD;
            btnApply.IsEnabled = false;
        }
        private void onApplyClicked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.USERNAME = SAGusername.Text;
            Properties.Settings.Default.PASSWORD = SAGpassword.Text;
            Properties.Settings.Default.Save();
            btnApply.IsEnabled = false;
        }
        private void onCloseClicked(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void usernameChanged(object sender, TextChangedEventArgs e)
        {
            btnApply.IsEnabled = true;
        }
        private void passwordChanged(object sender, TextChangedEventArgs e)
        {
            btnApply.IsEnabled = true;
        }
    }
}
