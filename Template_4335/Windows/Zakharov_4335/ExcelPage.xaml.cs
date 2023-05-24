using Microsoft.Win32;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4335.Windows.Zakharov_4335
{
    /// <summary>
    /// Логика взаимодействия для ExcelPage.xaml
    /// </summary>
    public partial class ExcelPage : Page
    {
        public ExcelPage()
        {
            InitializeComponent();
            InitializeComponent();
            DBGridModel.ItemsSource = IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => x.id_e).ToList();
        }
        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel|*.xlsx",
                Title = "Выберите файл"
            };
            if (!openFileDialog.ShowDialog() == true)
                return;
            ImportData(openFileDialog.FileName);
        }

        private static void ImportData(string path)
        {
            try
            {
                var objWorkExcel = new Excel.Application();
                var objWorkBook = objWorkExcel.Workbooks.Open(path);
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                var columns = lastCell.Column;
                var rows = lastCell.Row;
                var list = new string[rows, columns];
                for (var j = 0; j < columns; j++)
                {
                    for (var i = 1; i < rows; i++)
                    {
                        list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                objWorkBook.Close(false, Type.Missing, Type.Missing);
                objWorkExcel.Quit();
                GC.Collect();
                using (var db = new IsrpoEntities())
                {
                    for (var i = 1; i < 11; i++)
                    {
                        var uslugi = new employees
                        {
                            id_e = list[i, 0].ToString(),
                            role_e = list[i, 1].ToString(),
                            fio_e = list[i, 2].ToString(),
                            login_e = list[i, 3].ToString(),
                            password_e = list[i, 4].ToString(),
                            last_e = list[i, 5].ToString(),
                            type_e = list[i, 6].ToString(),
                        };
                        db.employees.Add(uslugi);
                    }
                    try
                    {
                        db.SaveChanges();
                        MessageBox.Show("Данные импортированы!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Внимание", MessageBoxButton.OK);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание1", MessageBoxButton.OK);
            }
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Объявление листов
            var first = new List<employees>();
            var second = new List<employees>();
            var third = new List<employees>();
            var categoriesPriceCount = 3;
            #endregion

            using (var isrpoEntities = new IsrpoEntities())
            {
                #region Сортировка
                first = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Продавец").OrderBy(fR => fR.id_e).ToList();
                second = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Администратор").OrderBy(fR => fR.id_e).ToList();
                third = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Старший смены").OrderBy(fR => fR.id_e).ToList();
                #endregion

                var app = new Excel.Application { SheetsInNewWorkbook = categoriesPriceCount };
                var book = app.Workbooks.Add(Type.Missing);

                #region Создание листов в Excel
                var startRowIndex = 1;
                var sheet1 = app.Worksheets.Item[1];
                sheet1.Name = "Продавцы";
                var sheet2 = app.Worksheets.Item[2];
                sheet2.Name = "Администраторы";
                var sheet3 = app.Worksheets.Item[3];
                sheet3.Name = "Старшие смены";
                #endregion

                #region Создание колонок в Excel
                sheet1.Cells[1][startRowIndex] = "Код сотрудника";
                sheet1.Cells[2][startRowIndex] = "ФИО";
                sheet1.Cells[3][startRowIndex] = "Логин";

                sheet2.Cells[1][startRowIndex] = "Код сотрудника";
                sheet2.Cells[2][startRowIndex] = "ФИО";
                sheet2.Cells[3][startRowIndex] = "Логин";

                sheet3.Cells[1][startRowIndex] = "Код сотрудника";
                sheet3.Cells[2][startRowIndex] = "ФИО";
                sheet3.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;
                #endregion

                #region Заполнение первого листа
                for (var i = 0; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in first)
                    {
                        sheet1.Cells[1][startRowIndex] = item.id_e;
                        sheet1.Cells[2][startRowIndex] = item.fio_e;
                        sheet1.Cells[3][startRowIndex] = item.login_e;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение второго листа
                for (var i = 1; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in second)
                    {
                        sheet2.Cells[1][startRowIndex] = item.id_e;
                        sheet2.Cells[2][startRowIndex] = item.fio_e;
                        sheet2.Cells[3][startRowIndex] = item.login_e;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение третьего листа
                for (var i = 2; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in third)
                    {
                        sheet3.Cells[1][startRowIndex] = item.id_e;
                        sheet3.Cells[2][startRowIndex] = item.fio_e;
                        sheet3.Cells[3][startRowIndex] = item.login_e;
                        startRowIndex++;
                    }
                }
                #endregion

                app.Visible = true;
            }
        }
    }
}
