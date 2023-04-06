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
import org.checkerframework.checker.nullness.qual.Nullable;
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
    static TableCellAddress find(Sheet sheet, @Nullable Object value,
                                 int startRow, int endRow,
                                 int startColumn, int endColumn) {
        @Nullable Object expected;
        if (value instanceof Number) {
            expected = ((Number) value).doubleValue();  // excel store Numbers as doubles
        } else {
            expected = value;
        }
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
            for (@Nullable Cell cell : row) {
                if (cell != null) {
                    int column = cell.getColumnIndex();
                    if (startColumn <= column && column < endColumn) {
                        if (predicate.test(cell)) {
                            CellAddress address = cell.getAddress();
                            return TableCellAddress.of(address.getRow(), address.getColumn());
                        }
                    }
                }
            }
        }
        return NOT_FOUND;
    }

    static @Nullable Object getValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue(); // returns double
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return getCachedFormulaValue(cell);
            case ERROR:
                throw new RuntimeException("Ячейка содержит ошибку вычисления формулы: " +
                        FormulaError.forInt(cell.getErrorCellValue()));
            case BLANK:
            case _NONE:
                return null;
            default:
                throw new RuntimeException("Unexpected cell type: " + cell.getCellType());
        }
    }

    private static @Nullable Object getCachedFormulaValue(Cell cell) {
        switch (cell.getCachedFormulaResultType()) {
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return cell.getRichStringCellValue();
            case ERROR:
                throw new RuntimeException("Ячейка не содержит кешированный результат формулы: " +
                        FormulaError.forInt(cell.getErrorCellValue()));
            default:
                return null; // never should occur
        }
    }

    private static boolean equals(Cell cell, @Nullable Object expected) {
        switch (cell.getCellType()) {
            case BLANK:
                return (expected == null) || Objects.equals(expected, "");
            case STRING:
                return (expected instanceof CharSequence) &&
                        Objects.equals(cell.getStringCellValue(), String.valueOf(expected));
            case NUMERIC:
                return (expected instanceof Number) &&
                        Math.abs((cell.getNumericCellValue() - ((Number) expected).doubleValue())) < 1e-6;
            case BOOLEAN:
                return Objects.equals(expected, cell.getBooleanCellValue());
            default:
                return false;
        }
    }
}
