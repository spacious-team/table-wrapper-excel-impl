/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2020  Vitalii Ananev <an-vitek@ya.ru>
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

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.spacious_team.table_wrapper.api.TableCell;
import org.spacious_team.table_wrapper.api.TableColumnDescription;
import org.spacious_team.table_wrapper.api.TableRow;

import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDateTime;

@RequiredArgsConstructor
public class ExcelTableCell implements TableCell {

    @Getter
    private final Cell cell;

    @Override
    public int getColumnIndex() {
        return cell.getColumnIndex();
    }

    @Override
    public Object getValue() {
        return ExcelTableHelper.getCellValue(cell);
    }

    @Override
    public int getIntValue() {
        return (int) getLongValue();
    }

    @Override
    public long getLongValue() {
        return ExcelTableHelper.getLongCellValue(cell);
    }

    @Override
    public Double getDoubleValue() {
        return ExcelTableHelper.getDoubleCellValue(cell);
    }

    @Override
    public BigDecimal getBigDecimalValue() {
        return ExcelTableHelper.getBigDecimalCellValue(cell);
    }

    @Override
    public String getStringValue() {
        return ExcelTableHelper.getStringCellValue(cell);
    }

    @Override
    public Instant getInstantValue() {
        return cell.getDateCellValue().toInstant();
    }

    @Override
    public LocalDateTime getLocalDateTimeValue() {
        return cell.getLocalDateTimeCellValue();
    }
}
