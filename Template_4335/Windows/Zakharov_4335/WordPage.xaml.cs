using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
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
using Word = Microsoft.Office.Interop.Word;

namespace Template_4335.Windows.Zakharov_4335
{
    /// <summary>
    /// Логика взаимодействия для WordPage.xaml
    /// </summary>
    public partial class WordPage : Page
    {
        public WordPage()
        {
            InitializeComponent();
            InitializeComponent();
            DBGridModel.ItemsSource = IsrpoEntities.GetContext().employees.AsEnumerable().OrderBy(x => x.id_e).ToList();
        }

        private async void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            string path = "C:\\Users\\liina\\Desktop\\Импорт\\4.json";
            using (var db = new IsrpoEntities())
            {
                try
                {
                    var emp = await JsonSerializer.DeserializeAsync<List<employees>>(new FileStream(path, FileMode.Open));
                    foreach (employees item in emp)
                    {
                        var employee = new employees
                        {
                            id_e = item.id_e,
                            role_e = item.role_e,
                            fio_e = item.fio_e,
                            login_e = item.login_e,
                            password_e = item.password_e,
                            last_e = item.last_e,
                            type_e = item.type_e
                        };

                        db.employees.Add(employee);
                    }
                    try
                    {
                        db.SaveChanges();
                        MessageBox.Show("Данные импортированы!");
                    }
                    catch (DbEntityValidationException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                catch { }
            }
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Объявление листов
            var first = new List<employees>();
            var second = new List<employees>();
            var third = new List<employees>();
            #endregion

            using (var isrpoEntities = new IsrpoEntities())
            {
                #region Сортировка
                first = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Продавец").OrderBy(fR => fR.id_e).ToList();
                second = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Администратор").OrderBy(fR => fR.id_e).ToList();
                third = isrpoEntities.employees.ToList().Where(sR => sR.role_e == "Старший смены").OrderBy(fR => fR.id_e).ToList();
                #endregion

                var app = new Word.Application();
                var document = app.Documents.Add();

                #region Заполнение первой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Продавцы (" + first.Count() + " сотрудника(-ов))";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, first.Count() + 1, 3);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код сотрудника";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Логин";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in first)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.id_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.fio_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение второй таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Администраторы (" + second.Count() + " сотрудника(-ов))";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, second.Count() + 1, 3);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код сотрудника";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Логин";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in second)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.id_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.fio_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение третьей таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Старшие смены (" + third.Count() + " сотрудника(-ов))";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, third.Count() + 1, 3);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код сотрудника";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Логин";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in third)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.id_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.fio_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.login_e;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                var paragraph1 = document.Paragraphs.Add();
                var range1 = paragraph1.Range;
                int a = first.Count() + second.Count() + third.Count();
                range1.Text = "Итого: " + a + " сотрудника(-ов)";
                paragraph1.set_Style("Заголовок 1");
                range1.InsertParagraphAfter();

                app.Visible = true;
            }
        }
    }
}
