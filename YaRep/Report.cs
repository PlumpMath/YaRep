using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YaRep
{
    /// <summary>
    /// Класс отчета
    /// </summary>
    public class Report
    {
        /// <summary>
        /// Конструктор приватный! Вместо этого используйте Report.Begin()   ;)
        /// </summary>
        private Report()
        {
        }

        /// <summary>
        /// Создает новый отчет
        /// </summary>
        /// <returns>экземпляр класса для работы с отчетом</returns>
        public static Report Begin()
        {
            return new Report();
        }

        /// <summary>
        /// Вызывается в конце работы над отчетом. В этот момент данные непосредственно выгружаются в excel и книга показывается пользователю.
        /// </summary>
        public void End()
        {
        }

        private Dictionary<string,Sheet> sheets = new Dictionary<string,Sheet>();
        public RandomAccessSheet CreateRandomAccessSheet(string sheetName)
        {
            var sheet = new RandomAccessSheet();
            sheets.Add(sheetName, sheet);
            return sheet;
        }
    }
}
