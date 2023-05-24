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

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Zakharov4335.xaml
    /// </summary>
    public partial class Zakharov4335 : Window
    {
        public Zakharov4335()
        {
            InitializeComponent();
        }

        private void WordPageBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new Windows.Zakharov_4335.ExcelPage());
        }

        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
