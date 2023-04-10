![java-version](https://img.shields.io/badge/Java-11-brightgreen?style=flat-square)
![jitpack-last-release](https://jitpack.io/v/spacious-team/table-wrapper-excel-impl.svg?style=flat-square)
[![Unit tests](https://img.shields.io/endpoint.svg?url=https%3A%2F%2Factions-badge.atrox.dev%2Fspacious-team%2Ftable-wrapper-excel-impl%2Fbadge%3Fref%3Ddevelop&style=flat-square&label=Test&logo=none)](
https://github.com/spacious-team/table-wrapper-excel-impl/actions/workflows/unit-tests.yml)
[![Coverage](https://img.shields.io/codecov/c/github/spacious-team/table-wrapper-excel-impl/develop?label=Coverage&style=flat-square&token=5DMBQMR6B0)](
https://codecov.io/gh/spacious-team/table-wrapper-excel-impl)

### Назначение
Предоставляет реализацию `Table Wrapper API` для удобного доступа к табличным данным, сохраненным в файлах формата
Microsoft Office Excel (xls) и [Office Open XML](https://ru.wikipedia.org/wiki/Office_Open_XML) (xlsx).

Если не используется Spring Boot Starter, то сначала в реестр доступных фабрик нужно зарегистрировать фабрику excel таблиц:
```java
TableFactoryRegistry.add(new ExcelTableFactory());
```

Пример создания таблиц с первого листа файла `1.xlsx`
```java
Workbook book = new XSSFWorkbook(Files.newInputStream(Path.of("1.xlsx")));
ReportPage reportPage = new ExcelSheet(book.getSheetAt(0));

Table table1 = reportPage.create("Table 1 description", ...);
...
Table tableN = reportPage.create("Table N description", ...);
```
Объекты `table`...`tableN` используются для удобного доступа к строкам и к значениям ячеек.

Больше подробностей в документации [Table Wrapper API](https://github.com/spacious-team/table-wrapper-api).

### Как использовать в своем проекте
Необходимо подключить репозиторий open source библиотек github [jitpack](https://jitpack.io/#spacious-team/table-wrapper-excel-impl),
например для Apache Maven проекта
```xml
<repositories>
    <repository>
        <id>central</id>
        <name>Central Repository</name>
        <url>https://repo.maven.apache.org/maven2</url>
        <snapshots>
            <enabled>false</enabled>
        </snapshots>
    </repository>
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
