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

        private string PrepareSheetNameForExcel(string sheetName)
        {
            return sheetName.Replace('.', '#');
        }

        /// <summary>
        /// Вызывается в конце работы над отчетом. В этот момент данные непосредственно выгружаются в excel и книга показывается пользователю.
        /// </summary>
        public void End()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            var workbook = excel.Workbooks.Add(Type.Missing);
            foreach (var sheet in sheets)
            {
                var excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelSheet.Name = PrepareSheetNameForExcel(sheet.Key);
                var data = sheet.Value.GetArray();
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                if ((cols < 1) || (rows < 1)) continue;
                excelSheet.get_Range(excelSheet.Cells[1, 1], excelSheet.Cells[rows, cols]).Value2 = data;
            }
            excel.Visible = true;
        }

        private Dictionary<string,Sheet> sheets = new Dictionary<string,Sheet>();
        public RandomAccessSheet CreateRandomAccessSheet(string sheetName)
        {
            var sheet = new RandomAccessSheet();
            sheets.Add(sheetName, sheet);
            return sheet;
        }

        public TableSheet CreateTableSheet(string sheetName)
        {
            var sheet = new TableSheet();
            sheets.Add(sheetName, sheet);
            return sheet;
        }
    }
}
