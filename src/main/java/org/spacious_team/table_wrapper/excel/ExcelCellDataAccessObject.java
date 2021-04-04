/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2021  Vitalii Ananev <an-vitek@ya.ru>
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
import org.spacious_team.table_wrapper.api.CellDataAccessObject;

import java.time.Instant;
import java.time.LocalDateTime;

public class ExcelCellDataAccessObject implements CellDataAccessObject<Cell, ExcelTableRow> {
    public static final ExcelCellDataAccessObject INSTANCE = new ExcelCellDataAccessObject();

    @Override
    public Cell getCell(ExcelTableRow row, Integer cellIndex) {
        return row.getRow().getCell(cellIndex);
    }

    @Override
    public Object getValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue(); // return double
            case BLANK -> null;
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> getCachedFormulaValue(cell);
            case ERROR -> throw new RuntimeException("Ячейка содержит ошибку вычисления формулы: " +
                    FormulaError.forInt(cell.getErrorCellValue()));
            case _NONE -> null;
        };
    }

    @Override
    public Instant getInstantValue(Cell cell) {
        return cell.getDateCellValue().toInstant();
    }

    @Override
    public LocalDateTime getLocalDateTimeValue(Cell cell) {
        return cell.getLocalDateTimeCellValue();
    }

    private static Object getCachedFormulaValue(Cell cell) {
        return switch (cell.getCachedFormulaResultType()) {
            case BOOLEAN -> cell.getBooleanCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> cell.getRichStringCellValue();
            case ERROR -> throw new RuntimeException("Ячейка не содержит кешированный результат формулы: " +
                    FormulaError.forInt(cell.getErrorCellValue()));
            default -> null; //never should occur
        };
    }
}
