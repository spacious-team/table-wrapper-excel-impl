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
import org.apache.poi.ss.usermodel.Row;
import org.spacious_team.table_wrapper.api.AbstractReportPageRow;
import org.spacious_team.table_wrapper.api.TableCell;

import java.util.Iterator;

import static org.spacious_team.table_wrapper.api.TableCellAddress.NOT_FOUND;


@RequiredArgsConstructor
public class ExcelTableRow extends AbstractReportPageRow {

    @Getter
    private final Row row;

    public TableCell getCell(int i) {
        Cell cell = row.getCell(i);
        return (cell == null) ? null : new ExcelTableCell(cell);
    }

    @Override
    public int getRowNum() {
        return row.getRowNum();
    }

    @Override
    public int getFirstCellNum() {
        return row.getFirstCellNum();
    }

    @Override
    public int getLastCellNum() {
        short lastCellNum = row.getLastCellNum(); // Gets the index of the last cell contained in this row PLUS ONE
        return (lastCellNum < 0) ? -1 : (lastCellNum - 1);
    }

    public boolean rowContains(Object value) {
        return ExcelTableHelper.find(row.getSheet(), value, row.getRowNum(), row.getRowNum() + 1,
                0, Integer.MAX_VALUE, String::equals) != NOT_FOUND;
    }

    @Override
    public Iterator<TableCell> iterator() {
        return new ReportPageRowIterator<>(row.iterator(), ExcelTableCell::new);
    }
}