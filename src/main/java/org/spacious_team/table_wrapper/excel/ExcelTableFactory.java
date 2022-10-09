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

import org.spacious_team.table_wrapper.api.AbstractTableFactory;
import org.spacious_team.table_wrapper.api.ReportPage;
import org.spacious_team.table_wrapper.api.Table;
import org.spacious_team.table_wrapper.api.TableCellRange;
import org.spacious_team.table_wrapper.api.TableColumnDescription;

@SuppressWarnings("unused")
public class ExcelTableFactory extends AbstractTableFactory<ExcelSheet> {

    public ExcelTableFactory() {
        super(ExcelSheet.class);
    }

    @Override
    public Table create(ReportPage reportPage,
                        String tableName,
                        TableCellRange tableRange,
                        Class<? extends TableColumnDescription> headerDescription,
                        int headersRowCount) {
        return new ExcelTable(
                cast(reportPage),
                tableName,
                tableRange,
                headerDescription,
                headersRowCount);
    }
}
