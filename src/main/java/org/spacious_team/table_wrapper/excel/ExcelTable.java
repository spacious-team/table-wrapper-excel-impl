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

import lombok.AccessLevel;
import lombok.Getter;
import lombok.ToString;
import org.spacious_team.table_wrapper.api.AbstractReportPage;
import org.spacious_team.table_wrapper.api.AbstractTable;
import org.spacious_team.table_wrapper.api.CellDataAccessObject;
import org.spacious_team.table_wrapper.api.TableCellRange;
import org.spacious_team.table_wrapper.api.TableColumnDescription;

@ToString(callSuper = true)
public class ExcelTable extends AbstractTable<ExcelTableRow> {

    @Getter(AccessLevel.PROTECTED)
    private final CellDataAccessObject<?, ExcelTableRow> cellDataAccessObject = ExcelCellDataAccessObject.INSTANCE;

    ExcelTable(AbstractReportPage<ExcelTableRow> reportPage,
               String tableName,
               TableCellRange tableRange,
               Class<? extends TableColumnDescription> headerDescription,
               int headersRowCount) {
        super(reportPage, tableName, tableRange, headerDescription, headersRowCount);
    }
}
