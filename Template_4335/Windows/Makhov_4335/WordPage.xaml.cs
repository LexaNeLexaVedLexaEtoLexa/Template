using Makhov_4335;
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
using Path = System.IO.Path;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4335.Windows.Makhov_4335
{
    /// <summary>
    /// Логика взаимодействия для WordPage.xaml
    /// </summary>
    public partial class WordPage : Page
    {
        public WordPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = SecondaryFunctions.Services;
        }

        /// <summary>
        /// Реализация импорта данных с помощью JSON формата.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportBtn_Click(object sender, System.Windows.RoutedEventArgs e) => ImportJsonData();

        /// <summary>
        /// Реализация экспорта из Базы Данных в Word.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportBtn_Click(object sender, System.Windows.RoutedEventArgs e) => ExportData();

        #region Методы импорта и экспорта
        //
        /// <summary>
        /// Импорт данных.
        /// </summary>
        private static async void ImportJsonData()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "1.json");

            using (var fileStream = new FileStream(path, FileMode.Open))
            using (var db = new WordExcelDatabaseEntities())
            {
                var services = await JsonSerializer.DeserializeAsync<List<Services>>(fileStream);

                foreach (Services item in services)
                {
                    var service = new Services
                    {
                        Id = item.Id,
                        ServiceName = item.ServiceName,
                        ServiceType = item.ServiceType,
                        ServiceCode = item.ServiceCode,
                        ServicePrice = item.ServicePrice
                    };

                    db.Services.Add(service);
                }
                try
                {
                    db.SaveChanges();
                    SecondaryFunctions.RefreshData();
                    MessageBox.Show("Данные импортированы успешно!",
                                    "Внимание!",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Information);
                    db.SaveChanges();
                }
                catch (DbEntityValidationException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// Экспорт данных.
        /// </summary>
        private static void ExportData()
        {
            #region Категории по возрасту

            var firstRangePrice = new List<Services>();
            var secondRangePrice = new List<Services>();
            var thirdRangePrice = new List<Services>();

            #endregion

            using (var db = new WordExcelDatabaseEntities())
            {
                #region Сортировка по цене

                firstRangePrice = db.Services.ToList()
                                             .Where(fR => Convert.ToInt32(fR.ServicePrice) >= 0 &&
                                                          Convert.ToInt32(fR.ServicePrice) <= 350)
                                             .GroupBy(fR => fR.ServiceName)
                                             .SelectMany(fR => fR)
                                             .ToList();

                secondRangePrice = db.Services.ToList()
                                              .Where(sR => Convert.ToInt32(sR.ServicePrice) >= 250 &&
                                                         Convert.ToInt32(sR.ServicePrice) <= 800)
                                              .GroupBy(sR => sR.ServiceName)
                                              .SelectMany(sR => sR)
                                              .ToList();

                thirdRangePrice = db.Services.ToList()
                                             .Where(tR => Convert.ToInt32(tR.ServicePrice) >= 800)
                                             .GroupBy(tR => tR.ServiceName)
                                             .SelectMany(tR => tR)
                                             .ToList();

                #endregion

                #region Создание Word 

                var app = new Word.Application();
                var document = app.Documents.Add();

                #endregion

                #region Создание параграфов

                #region Создание таблицы для первой категории

                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цен от 0 до 350 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, firstRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in firstRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.ServiceName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.ServiceType;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.ServicePrice;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                #region Создание таблицы для второй категории

                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цены от 250 до 800 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, secondRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in secondRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.ServiceName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.ServiceType;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.ServicePrice;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                #region Создание таблицы для третьей категории

                for (int i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория цен от 800 рублей";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var priceCategories = document.Tables.Add(tableRange, thirdRangePrice.Count() + 1, 4);

                    priceCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    priceCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    priceCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = priceCategories.Cell(1, 1).Range;
                    cellRange.Text = "Идентификатор";
                    cellRange = priceCategories.Cell(1, 2).Range;
                    cellRange.Text = "Наименование услуги";
                    cellRange = priceCategories.Cell(1, 3).Range;
                    cellRange.Text = "Тип услуги";
                    cellRange = priceCategories.Cell(1, 4).Range;
                    cellRange.Text = "Цена услуги";
                    priceCategories.Rows[1].Range.Bold = 1;
                    priceCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in thirdRangePrice)
                    {
                        cellRange = priceCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.ServiceName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.ServiceType;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = priceCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.ServicePrice;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                #endregion

                app.Visible = true;

                #endregion
            }
        }

        #endregion
    }
}
