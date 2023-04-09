/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2023  Spacious Team <spacious-team@ya.ru>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package org.spacious_team.table_wrapper.excel;

import nl.jqno.equalsverifier.EqualsVerifier;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.checkerframework.checker.nullness.qual.Nullable;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.spacious_team.table_wrapper.api.TableCell;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.Spliterators;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

import static java.time.ZoneOffset.UTC;
import static nl.jqno.equalsverifier.Warning.STRICT_INHERITANCE;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class ExcelTableRowTest {

    static Workbook workbook = new XSSFWorkbook();
    Row wrappeeRow;
    ExcelTableRow row;

    @BeforeEach
    void beforeEach() {
        Row sheetRow = workbook.createSheet()
                .createRow(10);
        wrappeeRow = spy(sheetRow);
        row = ExcelTableRow.of(wrappeeRow);
    }

    @AfterAll
    static void afterAll() throws IOException {
        workbook.close();
    }

    @Test
    void getRow() {
        assertSame(wrappeeRow, row.getRow());
    }

    @Test
    void getCell_Null() {
        @Nullable TableCell cell = row.getCell(0);

        assertNull(cell);
        verify(wrappeeRow).getCell(0);
    }

    @Test
    void getCell_nonNull() {
        Cell wrappeeCell = wrappeeRow.createCell(0);

        @Nullable TableCell cell = row.getCell(0);

        assertEquals(ExcelTableCell.of(wrappeeCell), cell);
        verify(wrappeeRow).getCell(0);
    }

    @Test
    void getRowNum() {
        int rowNum = row.getRowNum();

        assertEquals(10, rowNum);
        verify(wrappeeRow).getRowNum();
    }

    @Test
    void getFirstCellNum_noCells() {
        assertEquals(-1, row.getFirstCellNum());
    }

    @Test
    void getFirstCellNum_hasCells() {
        wrappeeRow.createCell(4);

        int firstCellNum = row.getFirstCellNum();

        assertEquals(4, firstCellNum);
        verify(wrappeeRow).getFirstCellNum();
    }

    @Test
    void getLastCellNum_noCells() {
        assertEquals(-1, row.getLastCellNum());
    }

    @Test
    void getLastCellNum_hasCells() {
        wrappeeRow.createCell(8);
        wrappeeRow.createCell(9);
        verify(wrappeeRow, times(2)).getLastCellNum();  // 2 times called on row create

        int lastCellNum = row.getLastCellNum();

        assertEquals(9, lastCellNum);
        verify(wrappeeRow, times(3)).getLastCellNum();  // 2 times called on row create
    }

    @Test
    void rowContains() {
        LocalDate localDate = LocalDate.of(2023, 4, 9);
        LocalDateTime localDateTime = LocalDateTime.of(localDate, LocalTime.of(17, 36, 1));
        Instant instant = localDateTime.atZone(UTC)
                .toInstant();
        Date date = Date.from(instant);
        createCells(localDate, localDateTime, date);

        assertTrue(row.rowContains(null));
        assertTrue(row.rowContains("test"));
        assertTrue(row.rowContains(true));
        assertTrue(row.rowContains(1));
        assertTrue(row.rowContains(1L));
        assertTrue(row.rowContains(2));
        assertTrue(row.rowContains(2L));
        assertTrue(row.rowContains(3.1f));
        assertTrue(row.rowContains(3.1));
        assertTrue(row.rowContains(3.2f));
        assertTrue(row.rowContains(3.2));
        assertTrue(row.rowContains(4));
        assertTrue(row.rowContains(5));
        assertTrue(row.rowContains(localDate));
        assertTrue(row.rowContains(localDateTime));
        assertTrue(row.rowContains(date));
        assertTrue(row.rowContains(instant));

        assertFalse(row.rowContains("test2"));
        assertFalse(row.rowContains(false));
        assertFalse(row.rowContains(8));
        assertFalse(row.rowContains(BigDecimal.valueOf(9.1)));
        assertFalse(row.rowContains(BigInteger.valueOf(10)));
        assertFalse(row.rowContains(LocalDate.now()));
        assertFalse(row.rowContains(LocalDateTime.now()));
        assertFalse(row.rowContains(Instant.now()));
        assertFalse(row.rowContains(Date.from(Instant.now())));
        assertFalse(row.rowContains(ZonedDateTime.ofInstant(instant, ZoneId.of("Europe/Paris"))));
        assertFalse(row.rowContains(OffsetDateTime.ofInstant(instant, ZoneId.of("Europe/Paris"))));
    }

    void createCells(LocalDate localDate, LocalDateTime localDateTime, Date date) {
        int i = 0;
        wrappeeRow.createCell(i++);  // null value
        wrappeeRow.createCell(i++).setCellValue("test");
        wrappeeRow.createCell(i++).setCellValue(true);
        wrappeeRow.createCell(i++).setCellValue(1);
        wrappeeRow.createCell(i++).setCellValue(2L);
        wrappeeRow.createCell(i++).setCellValue(3.1f);
        wrappeeRow.createCell(i++).setCellValue(3.2);
        wrappeeRow.createCell(i++).setCellValue((byte) 4);
        wrappeeRow.createCell(i++).setCellValue((short) 5);
        wrappeeRow.createCell(i++).setCellValue(localDate);
        wrappeeRow.createCell(i++).setCellValue(localDateTime);
        wrappeeRow.createCell(i).setCellValue(date);
    }

    @Test
    void iterator() {
        LocalDate localDate = LocalDate.of(2023, 4, 9);
        LocalDateTime localDateTime = LocalDateTime.of(localDate, LocalTime.of(17, 36, 1));
        Instant instant = localDateTime.atZone(UTC)
                .toInstant();
        Date date = Date.from(instant);
        createCells(localDate, localDateTime, date);
        Collection<@Nullable ExcelTableCell> expected = StreamSupport.stream(
                        Spliterators.spliteratorUnknownSize(wrappeeRow.iterator(), 0), false)
                .filter(Objects::nonNull)
                .map(ExcelTableCell::of)
                .collect(Collectors.toList());

        List<TableCell> actual = StreamSupport.stream(
                        Spliterators.spliteratorUnknownSize(row.iterator(), 0), false)
                .collect(Collectors.toList());

        assertEquals(expected, actual);
    }

    @Test
    void testEqualsAndHashCode() {
        EqualsVerifier
                .forClass(ExcelTableRow.class)
                .suppress(STRICT_INHERITANCE) // no subclass for test
                .verify();
    }

    @Test
    void testToString() {
        wrappeeRow.createCell(0).setCellValue("test");
        wrappeeRow.createCell(1).setCellValue(1);

        assertEquals("ExcelTableRow(rowIndex=10, firsrColumnIndex=0, lastColumnIndex=1)", row.toString());
    }
}