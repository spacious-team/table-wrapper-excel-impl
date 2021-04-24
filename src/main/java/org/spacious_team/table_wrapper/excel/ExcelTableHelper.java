/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2020  Vitalii Ananev <spacious-team@ya.ru>
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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.spacious_team.table_wrapper.api.TableCellAddress;

import java.util.function.BiPredicate;
import java.util.function.Predicate;

import static org.spacious_team.table_wrapper.api.TableCellAddress.NOT_FOUND;

class ExcelTableHelper {

    /**
     * @param value searching value
     * @param startRow search rows start from this
     * @param endRow search rows excluding this, can handle values greater than real rows count
     * @param startColumn search columns start from this
     * @param endColumn search columns excluding this, can handle values greater than real columns count
     * @return table table cell address or {@link TableCellAddress#NOT_FOUND}
     */
    static TableCellAddress find(Sheet sheet, Object value, int startRow, int endRow,
                                 int startColumn, int endColumn,
                                 BiPredicate<String, Object> stringPredicate) {
        if (sheet.getLastRowNum() == -1) {
            return NOT_FOUND;
        } else if (endRow > sheet.getLastRowNum()) {
            endRow = sheet.getLastRowNum() + 1; // endRow is exclusive
        }
        CellType type = getType(value);
        if (type == CellType.NUMERIC) {
            value = ((Number) value).doubleValue();
        }
        for(int rowNum = startRow; rowNum < endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;
            for (Cell cell : row) {
                if (cell != null) {
                    int column = cell.getColumnIndex();
                    if (startColumn <= column && column < endColumn && cell.getCellType() == type) {
                        if (compare(value, cell, stringPredicate)) {
                            CellAddress address = cell.getAddress();
                            return new TableCellAddress(address.getRow(), address.getColumn());
                        }
                    }
                }
            }
        }
        return NOT_FOUND;
    }

    private static TableCellAddress findByPredicate(Sheet sheet, int startRow, Predicate<Cell> predicate) {
        int endRow = sheet.getLastRowNum() + 1;
        for(int rowNum = startRow; rowNum < endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;
            for (Cell cell : row) {
                if (predicate.test(cell)) {
                    CellAddress address = cell.getAddress();
                    return new TableCellAddress(address.getRow(), address.getColumn());
                }
            }
        }
        return NOT_FOUND;
    }

    private static CellType getType(Object value) {
        CellType type;
        if (value instanceof String) {
            type = (((String) value).isEmpty()) ? CellType.BLANK : CellType.STRING;
        } else if (value instanceof Number) {
            type = CellType.NUMERIC;
        } else if (value instanceof Boolean) {
            type = CellType.BOOLEAN;
        } else if (value == null) {
            type = CellType.BLANK;
        } else {
            throw new IllegalArgumentException("Не могу сравнить значение '" + value + "' типа " + value.getClass().getName());
        }
        return type;
    }

    private static boolean compare(Object value, Cell cell, BiPredicate<String, Object> stringPredicate) {
        return switch (cell.getCellType()) {
            case BLANK -> value == null || value.equals("");
            case STRING -> stringPredicate.test(cell.getStringCellValue(), value);
            case NUMERIC -> (value instanceof Number) && (cell.getNumericCellValue() - ((Number) value).doubleValue()) < 1e-6;
            case BOOLEAN -> value.equals(cell.getBooleanCellValue());
            default -> false;
        };
    }
}
