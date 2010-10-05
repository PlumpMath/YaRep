using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YaRep
{
    /// <summary>
    /// Отчет со свободныи доступом в любую ячейку.
    /// </summary>
    public class RandomAccessSheet: Sheet
    {
        internal RandomAccessSheet()
        {
        }

        internal override object[,] GetArray()
        {
            object[,] result = new object[rows, columns];

            for (int rowNum = 0; rowNum < data.Length; rowNum++)
            {
                var row = data[rowNum];
                for (int colNum = 0; colNum < row.Length; colNum++)
                {
                    result[rowNum, colNum] = row[colNum];
                }
                for (int colNum = row.Length; colNum < columns; colNum++)
                {
                    result[rowNum, colNum] = defaultValue;
                }
            }

            for (int rowNum = data.Length; rowNum < rows; rowNum++)
            {
                for (int colNum = 0; colNum < columns; colNum++)
                {
                    result[rowNum, colNum] = defaultValue;
                }
            }

            return result;
        }

        private int rows = 0;
        private int columns = 0;
        private object[][] data = new object[0][];

        private void AllocRows(int newRows)
        {
            if (newRows<data.Length) return;
            object[][] newData = new object[newRows][];
            data.CopyTo(newData,0);
            for (int i=data.Length;i<newRows;i++)
            {
                newData[i] = new object[0];
            }
            data = newData;
            rows = newRows;
        }

        private object[] AllocCols(int newColumns, object[] rowArray)
        {
            if (newColumns < rowArray.Length) return rowArray;
            object[] newArray = new object[newColumns];
            rowArray.CopyTo(newArray, 0);
            for (int i = rowArray.Length; i < newColumns; i++)
            {
                newArray[i] = defaultValue;
            }
            if (newColumns > columns)
                columns = newColumns;
            return newArray;
        }

        /// <summary>
        /// Выделяет буфер под таблицу указанного размера.
        /// Если указанный размер меньше текущего, лишние данные обрезаются
        /// </summary>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        public void Resize(int rows, int columns)
        {
            object[][] newData = new object[rows][];
            for (int rownum = 0; rownum < data.Length; rownum++)
            {
                object[] row = data[rownum];
                int rowlen = row.Length;
                if (rowlen == columns)
                {
                    newData[rownum] = row;
                    continue;
                }
                object[] newRow = new object[columns];
                if (rowlen < columns)
                {
                    row.CopyTo(newRow, 0);
                }
                else
                {
                    for (int colnum = 0; colnum < columns; colnum++)
                    {
                        newRow[colnum] = row[colnum];
                    }
                }
                newData[rownum] = newRow;
            }
            data = newData;
        }


        private object defaultValue = null;
        public object DefaultValue { get {return defaultValue;} set{defaultValue = value;} }


        public object this[int row, int column]
        {
            get
            {
                if (row >= data.Length) return defaultValue;
                var row_array = data[row];
                if (column >= row_array.Length) return defaultValue;
                return row_array[column];
            }
            set
            {
                if (row >= rows)
                {
                    AllocRows(row+1);
                }
                var rowArray = data[row];
                if (column >= rowArray.Length)
                {
                    rowArray = AllocCols(column+1, rowArray);
                    data[row] = rowArray;
                }
                rowArray[column] = value;
            }
        }

        public bool IsEmpty
        {
            get
            {
                return ((rows==0)||(columns==0));
            }
        }


    }
}
