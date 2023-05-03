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

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.checkerframework.checker.nullness.qual.Nullable;
import org.spacious_team.table_wrapper.api.AbstractReportPage;
import org.spacious_team.table_wrapper.api.TableCellAddress;

import java.util.function.Predicate;

@RequiredArgsConstructor
public class ExcelSheet extends AbstractReportPage<ExcelTableRow> {

    @Getter
    private final Sheet sheet;

    @Override
    public TableCellAddress find(Object value, int startRow, int endRow, int startColumn, int endColumn) {
        return ExcelTableHelper.find(sheet, value, startRow, endRow, startColumn, endColumn);
    }

    @Override
    public TableCellAddress find(int startRow, int endRow, int startColumn, int endColumn,
                                 Predicate<@Nullable Object> cellValuePredicate) {
        return ExcelTableHelper.find(sheet, startRow, endRow, startColumn, endColumn,
                cell -> cellValuePredicate.test(ExcelTableHelper.getValue(cell)));
    }

    @Override
    public @Nullable ExcelTableRow getRow(int i) {
        Row row = sheet.getRow(i);
        return (row == null) ? null : ExcelTableRow.of(row);
    }

    @Override
    public int getLastRowNum() {
        return sheet.getLastRowNum();
    }

    /**
     * @param startRow first row for check
     * @return index of first empty row or -1 if not found
     */
    @Override
    public int findEmptyRow(int startRow) {
        int lastRowNum = startRow;
        for (int n = getLastRowNum(); lastRowNum <= n; lastRowNum++) {
            Row row = sheet.getRow(lastRowNum);
            if (row == null || row.getLastCellNum() == -1) {
                return lastRowNum;  // all row cells blank
            }
            boolean isEmptyRow = true;
            for (@Nullable Cell cell : row) {
                @Nullable Object value;
                if (!(cell == null
                        || ((value = ExcelCellDataAccessObject.INSTANCE.getValue(cell)) == null)
                        || (value instanceof String) && (value.toString().isEmpty()))) {
                    isEmptyRow = false;
                    break;
                }
            }
            if (isEmptyRow) {
                return lastRowNum;  // all row cells blank
            }
        }
        return -1;
    }
}
