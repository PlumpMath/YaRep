using System;
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
            object tst = ras[10,10];
            ras[1, 2] = 10;

            report.End();

        }
    }
}
