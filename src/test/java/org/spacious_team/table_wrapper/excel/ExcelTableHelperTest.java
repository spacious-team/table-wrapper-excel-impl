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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.spacious_team.table_wrapper.api.TableCellAddress;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class ExcelTableHelperTest {

    static Workbook workbook = new XSSFWorkbook();
    Sheet sheet;

    @BeforeEach
    void setUp() {
        sheet = workbook.createSheet();
    }

    @AfterAll
    static void afterAll() throws IOException {
        workbook.close();
    }

    @Test
    void find() {
        sheet.createRow(0).createCell(0).setCellValue("00");
        sheet.getRow(0).createCell(1).setCellValue("01");
        sheet.createRow(1).createCell(0).setCellValue("11");
        sheet.getRow(1).createCell(1).setCellValue(12);
        sheet.createRow(2).createCell(0);  // value == null
        sheet.getRow(2).createCell(1).setCellValue("22");


        assertEquals(TableCellAddress.of(1, 0),
                ExcelTableHelper.find(sheet, "11", 0, 3, 0, 2));
        assertEquals(TableCellAddress.of(1, 0),
                ExcelTableHelper.find(sheet, "11",
                        Integer.MIN_VALUE, Integer.MAX_VALUE, Integer.MIN_VALUE, Integer.MAX_VALUE));
        assertEquals(TableCellAddress.of(1, 1),
                ExcelTableHelper.find(sheet, 12, 0, 3, 0, 2));
        assertEquals(TableCellAddress.of(2, 0),
                ExcelTableHelper.find(sheet, null, 0, 3, 0, 2));
        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "00", 1, 3, 0, 2));
        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "00", 0, 3, 1, 2));
        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "00", -1, 0, 1, 2));
        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "00", 0, 3, -1, 0));
    }

    @Test
    void find_cellError() {
        sheet.createRow(0).createCell(0).setCellErrorValue((byte) 0);

        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "test cell with error", 0, 1, 0, 1));
    }

    @Test
    void find_exceptionallyFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("10/0");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertSame(TableCellAddress.NOT_FOUND,
                ExcelTableHelper.find(sheet, "test cell exception", 0, 1, 0, 1));
    }

    @Test
    void find_booleanFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("\"text1\"=\"text1\"");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals(TableCellAddress.of(0, 0),
                ExcelTableHelper.find(sheet, true, 0, 1, 0, 1));
    }

    @Test
    void find_numericFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("20+2");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals(TableCellAddress.of(0, 0),
                ExcelTableHelper.find(sheet, 22, 0, 1, 0, 1));
    }

    @Test
    void find_stringFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("concat(\"test\",\" string\")");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals(TableCellAddress.of(0, 0),
                ExcelTableHelper.find(sheet, "test string", 0, 1, 0, 1));
    }

    @Test
    void getValue_string() {
        String expected = "string";
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(expected);

        assertEquals(expected, ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_number() {
        double expected = 12;
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(expected);

        assertEquals(expected, ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_boolean() {
        boolean expected = true;
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(expected);

        assertEquals(expected, ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_blank() {
        Cell cell = sheet.createRow(0).createCell(0);
        assertNull(ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_exception() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellErrorValue((byte) 0);

        assertThrows(ArithmeticException.class, () -> ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_booleanFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("\"text1\"=\"text1\"");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals(true, ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_numericFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("20+2");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals(22.0, ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_stringFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("concat(\"test\",\" string\")");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertEquals("test string", ExcelTableHelper.getValue(cell));
    }

    @Test
    void getValue_exceptionallyFormula() {
        Cell cell = sheet.createRow(0).createCell(0);
        cell.setCellFormula("10/0");
        workbook.getCreationHelper()
                .createFormulaEvaluator()
                .evaluateFormulaCell(cell);

        assertThrows(ArithmeticException.class, () -> ExcelTableHelper.getValue(cell));
    }
}