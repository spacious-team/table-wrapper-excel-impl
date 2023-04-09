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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.util.Date;

import static nl.jqno.equalsverifier.Warning.STRICT_INHERITANCE;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class ExcelTableCellTest {

    static Workbook workbook = new XSSFWorkbook();
    Cell wrappeeCell;
    ExcelCellDataAccessObject dao;
    ExcelTableCell cell;

    @BeforeEach
    void setUp() {
        wrappeeCell = workbook.createSheet()
                .createRow(10)
                .createCell(20);
        dao = spy(ExcelCellDataAccessObject.INSTANCE);
        cell = ExcelTableCell.of(wrappeeCell, dao);
    }

    @AfterAll
    static void afterAll() throws IOException {
        workbook.close();
    }

    @Test
    void getColumnIndex() {
        assertEquals(20, cell.getColumnIndex());
    }

    @Test
    void getValue() {
        cell.getValue();
        verify(dao).getValue(wrappeeCell);
    }

    @ParameterizedTest
    @ValueSource(ints = {-1, 0, 2014})
    void getIntValue(int expected) {
        wrappeeCell.setCellValue(expected);
        assertEquals(expected, cell.getIntValue());
    }

    @ParameterizedTest
    @ValueSource(longs = {-1, 0, 2014})
    void getLongValue(long expected) {
        wrappeeCell.setCellValue(expected);
        assertEquals(expected, cell.getLongValue());
    }

    @ParameterizedTest
    @ValueSource(doubles = {0, 10.24, 10.24000, 10.2400000000001, 10.2400000000000000000000000000000000001})
    void getDoubleValue(double expected) {
        wrappeeCell.setCellValue(expected);
        assertEquals(expected, cell.getDoubleValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"0", "10.24", "10.24000000000001"})
    void getBigDecimalValue(String value) {
        BigDecimal expected = new BigDecimal(value);
        wrappeeCell.setCellValue(expected.doubleValue());
        assertEquals(expected, cell.getBigDecimalValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"10.24", "abc", "This is", "Это есть", "true", "0"})
    void getStringValue(String expected) {
        wrappeeCell.setCellValue(expected);
        assertEquals(expected, cell.getStringValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"2023-03-13T20:15:30Z", "2023-03-13T20:15:30.123Z"})
    void getInstantValue(String dateTime) {
        Instant expected = Instant.parse(dateTime);
        Date cellValue = Date.from(expected);
        wrappeeCell.setCellValue(cellValue);
        assertEquals(expected, cell.getInstantValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"2023-03-13T20:15:30.123456Z", "2023-03-13T20:15:30.123456789Z"})
    void getInstantValue_nanosLost(String dateTime) {
        Instant instant = Instant.parse(dateTime);
        Date cellValue = Date.from(instant);
        wrappeeCell.setCellValue(cellValue);
        // xml cell value lost nanos part, calc expected instant
        Instant expected = Instant.parse("2023-03-13T20:15:30.123Z");

        assertEquals(expected, cell.getInstantValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"2023-03-13T20:15:30Z", "2023-03-13T20:15:30.123Z"})
    void getLocalDateTimeValue(String dateTime) {
        Instant instant = Instant.parse(dateTime);
        Date cellValue = Date.from(instant);
        wrappeeCell.setCellValue(cellValue);
        LocalDateTime expected = instant.atZone(ZoneId.systemDefault())
                .toLocalDateTime();
        assertEquals(expected, cell.getLocalDateTimeValue());
    }

    @ParameterizedTest
    @ValueSource(strings = {"2023-03-13T20:15:30Z", "2023-03-13T20:15:30.123Z"})
    void getLocalDateTimeValue_withZoneId(String dateTime) {
        Instant instant = Instant.parse(dateTime);
        Date cellValue = Date.from(instant);
        wrappeeCell.setCellValue(cellValue);
        LocalDateTime expected = instant.atZone(ZoneId.systemDefault())
                .toLocalDateTime();
        int offsetSeconds = ZoneId.systemDefault()
                .getRules()
                .getOffset(expected)
                .getTotalSeconds();
        ZoneId zoneIdPlusHour = ZoneOffset.ofTotalSeconds(offsetSeconds + 3600);

        assertEquals(expected, cell.getLocalDateTimeValue());
        assertEquals(expected, cell.getLocalDateTimeValue(ZoneOffset.systemDefault()));
        assertEquals(expected.plusHours(1), cell.getLocalDateTimeValue(zoneIdPlusHour));
    }

    @Test
    void createWithCellDataAccessObject() {
        ExcelCellDataAccessObject dao = mock(ExcelCellDataAccessObject.class);

        ExcelTableCell actual = cell.createWithCellDataAccessObject(dao);

        assertNotSame(cell, actual);
        assertNotSame(dao, cell.getCellDataAccessObject());
        assertSame(dao, actual.getCellDataAccessObject());
    }

    @Test
    void testEqualsAndHashCode() {
        EqualsVerifier
                .forClass(ExcelTableCell.class)
                .suppress(STRICT_INHERITANCE) // no subclass for test
                .verify();
    }

    @Test
    void testToString() {
        wrappeeCell.setCellValue("data");
        assertEquals("ExcelTableCell(value=data)", cell.toString());
    }
}