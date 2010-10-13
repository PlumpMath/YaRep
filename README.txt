YaRep
=====

Yet Another Report library is a simple wrapper for excel interop.
YaRep is quite easy to use!



var report = YaRep.Report.Begin();

var table = report.CreateTableSheet("table sheet example");
table.AddColumns("col1", "col2");
table.AddRow("asdf", 123);
table.AddRow("fdsa", 456);

var ras = report.CreateRandomAccessSheet("randos access sheet example");
ras[2,15] = 123;
ras[55,55] = "asdf";
ras[5,5] = something;

report.End();