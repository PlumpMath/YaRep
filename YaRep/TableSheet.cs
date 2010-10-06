using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YaRep
{
    /// <summary>
    /// Отчет, представляющий из себя обыкновенную таблицу.
    /// Колонки ее именованы.
    /// Данные добавляются построчно.
    /// </summary>
    public class TableSheet : Sheet
    {
        

        /// <summary>
        /// Конструктор. Ничего не делает :]
        /// </summary>
        internal TableSheet()
        {
        }



        private RandomAccessSheet header = new RandomAccessSheet();
        /// <summary>
        /// Блок со свободной разметкой, который будет выведен перед таблицей
        /// </summary>
        public RandomAccessSheet Header
        {
            get
            {
                return header;
            }
        }

        private IList<string> columns = new List<string>();
        private IList<object[]> rows = null;



        internal override object[,] GetArray()
        {
            object[,] headerArray = Header.GetArray();
            if (columns.Count == 0) return headerArray;

            if (rows == null) rows = new List<object[]>();  // TODO: Подумать, как техничнее обойти ситуацию, когда записей нет... но шапку-то выводить надо!

            int headerRows = headerArray.GetLength(0);
            int headerCols = headerArray.GetLength(0);
            int resultCols = columns.Count;
            int resultRows = headerRows + 1 + rows.Count;

            object[,] result = new object[resultRows, Math.Max(resultCols,headerCols)];

            for (int rowNum = 0; rowNum < headerRows; rowNum++)
            {
                for (int colNum = 0; colNum < headerCols; colNum++)
                {
                    result[rowNum, colNum] = headerArray[rowNum, colNum];
                }
            }
            

            for (int columnNum = 0; columnNum < resultCols; columnNum++)
            {
                result[headerRows, columnNum] = columns[columnNum];
            }

            int currentRowNum = headerRows;
            foreach (var row in rows)
            {
                currentRowNum++;
                for (int col = 0; col < resultCols; col++)
                {
                    result[currentRowNum, col] = row[col];
                }
            }

            return result;
        }

        /// <summary>
        /// Добавляет колонку с заданным именем.
        /// Нельзя добавлять колонки, если уже были добавлены строки
        /// </summary>
        /// <param name="columnName">имя колонки</param>
        public void AddColumn(string columnName)
        {
            if (rows!=null) throw new Exception("Нельзя добавлять колонки, если уже вызывался метод AddRow");
            columns.Add(columnName);
        }
        /// <summary>
        /// Добавляет несколько колонок.
        /// Нельзя добавлять колонки, если уже были добавлены строки
        /// </summary>
        /// <param name="columns">имена колонок</param>
        public void AddColumns(params string[] columns)
        {
            foreach (var column in columns)
            {
                AddColumn(column);
            }
        }

        /// <summary>
        /// Добавляет строку.
        /// Количество параметров должно совпадать с количеством объявленных колонок
        /// </summary>
        /// <param name="row"></param>
        public void AddRow(params object[] row)
        {
            if (rows == null) rows = new List<object[]>();

            if (row.Length != columns.Count) throw new Exception("Метод AddRow получил набор параметров, не соответсвующй количеству колонок");

            rows.Add(row.Clone() as object[]);  // Клонируем массив! :]
        }
    }
}
