YaRep
=====

Yet Another Report library - простенькая библиотека для создания отчетов в Excel.

Библиотека использует Excel Interop.

Нужна только потому, что запись в interop "по ячейке" - ну очень медленная :]



Применяется библиотека примерно так:

var report = Report.Begin();
var ras = report.CreateRandomAccessSheet("randos access sheet example");
//ras.DefalutValues = "N/A";
ras[2,15] = 123;
ras[55,55] = "asdf";
ras[5,5] = something;
report.End();