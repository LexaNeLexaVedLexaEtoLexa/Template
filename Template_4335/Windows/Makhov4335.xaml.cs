using LaboratoryWorkUsingWordExcel.Classes;
using Makhov_4335;
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
using Template_4335.Windows.Makhov_4335;

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Zakharov4335.xaml
    /// </summary>
    public partial class Makhov4335 : Window
    {
        public Makhov4335()
        {
            InitializeComponent();
            MainFrame.Navigate(new ExcelPage());
            Manager.MainFrame = MainFrame;
        }
        /// <summary>
        /// Открыть страницу для взаимодействия с Excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExcelPageBtn_OnClick(object sender, RoutedEventArgs e) =>
            MainFrame.Navigate(new ExcelPage());

        /// <summary>
        /// Открыть страницу для взаимодействия с Word.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WordPageBtn_OnClick(object sender, RoutedEventArgs e) =>
            MainFrame.Navigate(new WordPage());

        /// <summary>
        /// Очистка таблицы из Базы Данных.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {
            if (
                MessageBox.Show("Вы уверены в том, что хотите очистить Базу Данных?",
                                "Внимание!",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Question) == MessageBoxResult.Yes
               )
            {
                SecondaryFunctions.DeleteData();
                SecondaryFunctions.RefreshData();
            }
        }

        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {

        }
        //private void WordPageBtn_Click(object sender, RoutedEventArgs e)
        //{
        //    MainFrame.Navigate(new Makhov_4335.WordPage());
        //}

        //private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        //{
        //    MainFrame.Navigate(new Makhov_4335.ExcelPage());
        //}

        //private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        //{
        //    if (MessageBox.Show("Очистить данные?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        //    {
        //        using (var isrpoEntities = new Makhov_4335.WordExcelDatabaseEntities())
        //        {
        //            isrpoEntities.Services.RemoveRange(isrpoEntities.Services.ToList());
        //            isrpoEntities.SaveChanges();
        //            Makhov_4335.WordExcelDatabaseEntities.GetContext().Services.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Clear();
        //            foreach (var uslugi in isrpoEntities.Services.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList())
        //            {
        //                Makhov_4335.WordExcelDatabaseEntities.GetContext().Services.AsEnumerable().OrderBy(x => Convert.ToInt32(x.id_e)).ToList().Add(uslugi);
        //            }
        //        }
        //    }
        //}
    }
}
