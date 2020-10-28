![java-version](https://img.shields.io/badge/Java-14-brightgreen?style=flat-square)
![jitpack-last-release](https://jitpack.io/v/spacious-team/table-wrapper-excel-impl.svg?style=flat-square)

### Назначение
Предоставляет реализацию `Table Wrapper API` для удобного доступа к табличным данным, сохраненным в файлах формата
Microsoft Office Excel (xls) и [Office Open XML](https://ru.wikipedia.org/wiki/Office_Open_XML) (xlsx).

Пример создания фабрики таблиц
```java
Workbook book = new XSSFWorkbook(Files.newInputStream(Path.get("1.xlsx")));
ReportPage reportPage = new ExcelSheet(book.getSheetAt(0));
TableFactory tableFactory = new ExcelTableFactory();

Table table1 = tableFactory.create(reportPage, "Table 1 description", ...);
...
Table tableN = tableFactory.create(reportPage, "Table N description", ...);
```
Объекты `table`...`tableN` используются для удобного доступа к строкам и к значениям ячеек.

Больше подробностей в документации [Table Wrapper API](https://github.com/spacious-team/table-wrapper-api).

### Как использовать в своем проекте
Необходимо подключить репозиторий open source библиотек github [jitpack](https://jitpack.io/#spacious-team/table-wrapper-excel-impl),
например для Apache Maven проекта
```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>
```
и добавить зависимость
```xml
<dependency>
    <groupId>com.github.spacious-team</groupId>
    <artifactId>table-wrapper-excel-impl</artifactId>
    <version>master-SNAPSHOT</version>
</dependency>
```
В качестве версии можно использовать:
- версию [релиза](https://github.com/spacious-team/table-wrapper-excel-impl/releases) на github;
- паттерн `<branch>-SNAPSHOT` для сборки зависимости с последнего коммита выбранной ветки;
- короткий 10-ти значный номер коммита для сборки зависимости с указанного коммита.
