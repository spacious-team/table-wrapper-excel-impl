/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2021  Spacious Team <spacious-team@ya.ru>
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
import org.checkerframework.checker.nullness.qual.Nullable;
import org.spacious_team.table_wrapper.api.CellDataAccessObject;

import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDateTime;

import static java.util.Objects.requireNonNull;

public class ExcelCellDataAccessObject implements CellDataAccessObject<Cell, ExcelTableRow> {
    public static final ExcelCellDataAccessObject INSTANCE = new ExcelCellDataAccessObject();

    @Override
    public @Nullable Cell getCell(ExcelTableRow row, Integer cellIndex) {
        return row.getRow().getCell(cellIndex);
    }

    @Override
    public @Nullable Object getValue(Cell cell) {
        return ExcelTableHelper.getValue(cell);
    }

    @Override
    public BigDecimal getBigDecimalValue(Cell cell) {
        double number = getDoubleValue(cell);
        return (Double.compare(number, 0D) == 0) ? BigDecimal.ZERO : BigDecimal.valueOf(number);
    }

    @Override
    public String getStringValue(Cell cell) {
        @SuppressWarnings({"nullness", "ConstantConditions"})
        Object value = requireNonNull(getValue(cell), "Not a string");
        String strValue = value.toString();
        if ((value instanceof Number) && strValue.endsWith(".0")) {
            return strValue.substring(0, strValue.length() - 2);
        }
        return strValue;
    }

    @Override
    public Instant getInstantValue(Cell cell) {
        return cell.getDateCellValue().toInstant();
    }

    @Override
    public LocalDateTime getLocalDateTimeValue(Cell cell) {
        return cell.getLocalDateTimeCellValue();
    }
}
