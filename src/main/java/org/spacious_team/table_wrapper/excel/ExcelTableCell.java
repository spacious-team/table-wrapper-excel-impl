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

import lombok.EqualsAndHashCode;
import lombok.ToString;
import org.apache.poi.ss.usermodel.Cell;
import org.spacious_team.table_wrapper.api.AbstractTableCell;

import static org.spacious_team.table_wrapper.excel.ExcelCellDataAccessObject.INSTANCE;

@ToString
@EqualsAndHashCode(callSuper = true)
public class ExcelTableCell extends AbstractTableCell<Cell, ExcelCellDataAccessObject> {

    public static ExcelTableCell of(Cell cell) {
        return of(cell, INSTANCE);
    }

    public static ExcelTableCell of(Cell cell, ExcelCellDataAccessObject dao) {
        return new ExcelTableCell(cell, dao);
    }

    private ExcelTableCell(Cell cell, ExcelCellDataAccessObject dao) {
        super(cell, dao);
    }

    @Override
    public int getColumnIndex() {
        return getCell().getColumnIndex();
    }

    @Override
    protected ExcelTableCell createWithCellDataAccessObject(ExcelCellDataAccessObject dao) {
        return new ExcelTableCell(getCell(), dao);
    }

    @SuppressWarnings("unused")
    @ToString.Include(name = "value")
    private String getCellData() {
        return getStringValue();
    }
}
