/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2020  Spacious Team <spacious-team@ya.ru>
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
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.spacious_team.table_wrapper.api.TableCellAddress;

import java.util.Objects;
import java.util.function.Predicate;

import static org.spacious_team.table_wrapper.api.TableCellAddress.NOT_FOUND;

class ExcelTableHelper {

    /**
     * @param value       searching value
     * @param startRow    search rows start from this
     * @param endRow      search rows excluding this, can handle values greater than real rows count
     * @param startColumn search columns start from this
     * @param endColumn   search columns excluding this, can handle values greater than real columns count
     * @return table cell address or {@link TableCellAddress#NOT_FOUND}
     */
    static TableCellAddress find(Sheet sheet, Object value,
                                 int startRow, int endRow,
                                 int startColumn, int endColumn) {
        Object expected = (value instanceof Number) ?
                ((Number) value).doubleValue() : // excel store Numbers as doubles
                value;
        return find(sheet, startRow, endRow, startColumn, endColumn, (cell) -> equals(cell, expected));
    }

    /**
     * @param startRow    search rows start from this
     * @param endRow      search rows excluding this, can handle values greater than real rows count
     * @param startColumn search columns start from this
     * @param endColumn   search columns excluding this, can handle values greater than real columns count
     * @return table cell address or {@link TableCellAddress#NOT_FOUND}
     */
    static TableCellAddress find(Sheet sheet, int startRow, int endRow,
                                 int startColumn, int endColumn,
                                 Predicate<Cell> predicate) {
        startRow = Math.max(0, startRow);
        endRow = Math.min(endRow, sheet.getLastRowNum() + 1); // endRow is exclusive
        for (int rowNum = startRow; rowNum < endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell != null) {
                    int column = cell.getColumnIndex();
                    if (startColumn <= column && column < endColumn) {
                        if (predicate.test(cell)) {
                            CellAddress address = cell.getAddress();
                            return new TableCellAddress(address.getRow(), address.getColumn());
                        }
                    }
                }
            }
        }
        return NOT_FOUND;
    }

    static Object getValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue(); // returns double
            case BLANK -> null;
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> getCachedFormulaValue(cell);
            case ERROR -> throw new RuntimeException("Ячейка содержит ошибку вычисления формулы: " +
                    FormulaError.forInt(cell.getErrorCellValue()));
            case _NONE -> null;
        };
    }

    private static Object getCachedFormulaValue(Cell cell) {
        return switch (cell.getCachedFormulaResultType()) {
            case BOOLEAN -> cell.getBooleanCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> cell.getRichStringCellValue();
            case ERROR -> throw new RuntimeException("Ячейка не содержит кешированный результат формулы: " +
                    FormulaError.forInt(cell.getErrorCellValue()));
            default -> null; // never should occur
        };
    }

    private static boolean equals(Cell cell, Object expected) {
        return switch (cell.getCellType()) {
            case BLANK -> expected == null || expected.equals("");
            case STRING -> (expected instanceof CharSequence) && Objects.equals(cell.getStringCellValue(), expected.toString());
            case NUMERIC -> (expected instanceof Number) && Math.abs((cell.getNumericCellValue() - ((Number) expected).doubleValue())) < 1e-6;
            case BOOLEAN -> expected.equals(cell.getBooleanCellValue());
            default -> false;
        };
    }
}
