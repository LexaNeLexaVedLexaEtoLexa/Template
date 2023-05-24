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
            MainFrame.Navigate(new Zakharov_4335.WordPage());
        }

        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new Zakharov_4335.ExcelPage());
        }

        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить данные?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                using (var isrpoEntities = new Zakharov_4335.IsrpoEntities())
                {
                    isrpoEntities.employees.RemoveRange(isrpoEntities.employees.ToList());
                    isrpoEntities.SaveChanges();
                    Zakharov_4335.IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Clear();
                    foreach (var uslugi in isrpoEntities.employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList())
                    {
                        Zakharov_4335.IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Add(uslugi);
                    }
                }
            }
        }
    }
}
