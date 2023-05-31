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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.checkerframework.checker.nullness.qual.Nullable;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.Test;
import org.spacious_team.table_wrapper.api.TableCellAddress;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.spy;
import static org.mockito.Mockito.verify;

class ExcelSheetTest {

    static Workbook workbook = new XSSFWorkbook();

    @AfterAll
    static void afterAll() throws IOException {
        workbook.close();
    }

    @Test
    void find() {
        Sheet worksheet = getTestSheet();
        ExcelSheet reportPage = new ExcelSheet(worksheet);

        assertEquals(TableCellAddress.of(0, 1),
                reportPage.find("12", 0, 2, 0, 2));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find("12", 1, 2, 0, 2));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find("12", 0, 2, 2, 3));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find("xyz", 0, 2, 0, 2));
    }

    @Test
    void findByPrefix() {
        Sheet worksheet = getTestSheet();
        ExcelSheet reportPage = new ExcelSheet(worksheet);

        assertEquals(TableCellAddress.of(0, 1),
                reportPage.find(0, 2, 0, 2, "12"::equals));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find(1, 2, 0, 2, "12"::equals));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find(0, 2, 2, 3, "12"::equals));
        assertEquals(TableCellAddress.NOT_FOUND,
                reportPage.find(0, 2, 0, 2, "xyz"::equals));
    }

    @Test
    void getRow() {
        int row = 1;
        Sheet sheet = spy(getTestSheet());
        ExcelSheet reportPage = new ExcelSheet(sheet);

        @Nullable ExcelTableRow actual = reportPage.getRow(row);

        assertInstanceOf(ExcelTableRow.class, actual);
        verify(sheet).getRow(row);
        assertNull(reportPage.getRow(2));
        assertNull(reportPage.getRow(-1));
    }

    @Test
    void getLastRowNum() {
        Sheet sheet = spy(getTestSheet());
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(1, reportPage.getLastRowNum());
    }

    @Test
    void getLastRowNumEmptyPage() {
        Sheet sheet = workbook.createSheet();
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(-1, reportPage.getLastRowNum());
    }

    @Test
    void findEmptyRow_noEmpty() {
        Sheet sheet = getTestSheet();
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(-1, reportPage.findEmptyRow(0));
    }

    @Test
    void findEmptyRow_onEmptySheet() {
        Sheet sheet = workbook.createSheet();
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(-1, reportPage.findEmptyRow(0));
    }

    @Test
    void findEmptyRow_onSheetOfEmptyRow() {
        Sheet sheet = workbook.createSheet();
        sheet.createRow(0);
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(0, reportPage.findEmptyRow(0));
    }

    @Test
    void findEmptyRow() {
        Sheet sheet = getTestSheet();
        sheet.createRow(2).createCell(0).setCellValue("");
        sheet.createRow(2).createCell(1).setCellValue("");
        ExcelSheet reportPage = new ExcelSheet(sheet);

        assertEquals(2, reportPage.findEmptyRow(0));
    }

    Sheet getTestSheet() {
        Sheet sheet = workbook.createSheet();
        sheet.createRow(0).createCell(0).setCellValue("11");
        sheet.createRow(0).createCell(1).setCellValue("12");
        sheet.createRow(1).createCell(0).setCellValue("21");
        sheet.createRow(1).createCell(1).setCellValue("22");
        return sheet;
    }
}