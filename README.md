![java-version](https://img.shields.io/badge/Java-14-brightgreen?style=flat-square)
![jitpack-last-release](https://jitpack.io/v/spacious-team/table-wrapper-excel-impl.svg?style=flat-square)

Реализация `Table Wrapper API` для таблиц, сохраненных в файлах формата Microsoft Office Excel (xls) и
[Office Open XML](https://ru.wikipedia.org/wiki/Office_Open_XML) (xlsx).

Пример создания фабрики таблиц
```java
Workbook book = new XSSFWorkbook(Files.newInputStream(Path.get("1.xlsx")));
ReportPage reportPage = new ExcelSheet(book.getSheetAt(0));
TableFactory tableFactory = new ExcelTableFactory(reportPage);

Table table1 = tableFactory.create(reportPage, "Table 1 description", ...);
...
Table tableN = tableFactory.create(reportPage, "Table N description", ...);
```
Таблицы `table`...`tableN` используются для удобного доступа к строкам и к значениям ячеек.

Больше подробностей в документации [Table Wrapper API](https://github.com/spacious-team/table-wrapper-api).
