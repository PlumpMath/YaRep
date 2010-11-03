using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

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
            if (sheets.Count < 1) return;
            Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            var workbook = excel.Workbooks.Add(Type.Missing);
            int origSheets = workbook.Sheets.Count;
            Worksheet[] origSheet = new Worksheet[origSheets];
            for(int i=0;i<origSheets;i++)
            {
                origSheet[i] = (Worksheet)workbook.Sheets.get_Item(i+1);
            }
            foreach (var sheet in sheets)
            {
                var data = sheet.Value.GetArray();
                if (data == null) continue;
                var excelSheet = (Worksheet)workbook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelSheet.Name = PrepareSheetNameForExcel(sheet.Key);
                int rows = data.GetLength(0);
                int cols = data.GetLength(1);
                if ((cols < 1) || (rows < 1)) continue;
                excelSheet.get_Range(excelSheet.Cells[1, 1], excelSheet.Cells[rows, cols]).Value2 = data;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
            }
            foreach (var sheet in origSheet)
            {
                sheet.Delete();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            }
            excel.Visible = true;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
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

        public AccumulationSheet CreateAccumulationSheet(string sheetName)
        {
            var sheet = new AccumulationSheet();
            sheets.Add(sheetName, sheet);
            return sheet;
        }
    }
}
