
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using Template_4335.Windows.Makhov_4335;

namespace Makhov_4335
{
    public static class SecondaryFunctions
    {
        /// <summary>
        /// Отслеживание изменений в таблице услуг.
        /// </summary>
        public static ObservableCollection<Services> Services
        {
            get;
            set;
        } = new ObservableCollection<Services> (WordExcelDatabaseEntities.GetContext()
                                                                         .Services
                                                                         .AsEnumerable()
                                                                         .OrderBy(x => Convert.ToInt32(x.Id))
                                                                         .ToList()
                                                );

        /// <summary>
        /// Очистка данных из Базы Данных.
        /// </summary>
        public static void DeleteData()
        {
            using (var db = new WordExcelDatabaseEntities())
            {
                db.Services.RemoveRange(db.Services.ToList());
                db.SaveChanges();
            }
        }
        
        /// <summary>
        /// Обновление данных в таблицею
        /// </summary>
        public static void RefreshData()
        {
            using (var db = new WordExcelDatabaseEntities())
            {
                var services = db.Services
                                 .AsEnumerable()
                                 .OrderBy(x => Convert.ToInt32(x.Id))
                                 .ToList();
                Services.Clear();
                foreach (var service in services)
                {
                    Services.Add(service);
                }
            }
        }
    }
}