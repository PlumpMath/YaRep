﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Example1
{
    class Program
    {
        static void Main(string[] args)
        {
            var report = YaRep.Report.Begin();


            var ras = report.CreateRandomAccessSheet("tst");
            ras[1, 2] = 10;


            var table = report.CreateTableSheet("table_test");
            table.Header[0, 0] = "Some header";
            table.AddColumns("col1", "col2");
            table.AddRow("asdf", 123);
            table.AddRow("fdsa",456);


            report.End();
        }
    }
}
