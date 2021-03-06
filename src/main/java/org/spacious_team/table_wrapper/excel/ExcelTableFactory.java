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

import org.spacious_team.table_wrapper.api.AbstractTableFactory;
import org.spacious_team.table_wrapper.api.ReportPage;
import org.spacious_team.table_wrapper.api.Table;
import org.spacious_team.table_wrapper.api.TableColumnDescription;

public class ExcelTableFactory extends AbstractTableFactory<ExcelSheet> {

    public ExcelTableFactory() {
        super(ExcelSheet.class);
    }

    public Table create(ReportPage reportPage,
                        String tableName,
                        String lastRowString,
                        Class<? extends TableColumnDescription> headerDescription,
                        int headersRowCount) {
        return new ExcelTable(cast(reportPage),
                tableName,
                reportPage.getTableCellRange(tableName, headersRowCount, lastRowString),
                headerDescription,
                headersRowCount);
    }

    @Override
    public Table create(ReportPage reportPage,
                        String tableName,
                        Class<? extends TableColumnDescription> headerDescription,
                        int headersRowCount) {
        return new ExcelTable(cast(reportPage),
                tableName,
                reportPage.getTableCellRange(tableName, headersRowCount),
                headerDescription,
                headersRowCount);
    }

    @Override
    public Table createNameless(ReportPage reportPage,
                                String providedTableName,
                                String firstRowString,
                                String lastRowString,
                                Class<? extends TableColumnDescription> headerDescription,
                                int headersRowCount) {
        return new ExcelTable(cast(reportPage),
                providedTableName,
                reportPage.getTableCellRange(firstRowString, headersRowCount, lastRowString)
                        .addRowsToTop(1),
                headerDescription,
                headersRowCount);
    }


    @Override
    public Table createNameless(ReportPage reportPage,
                                String providedTableName,
                                String firstRowString,
                                Class<? extends TableColumnDescription> headerDescription,
                                int headersRowCount) {
        return new ExcelTable(cast(reportPage),
                providedTableName,
                reportPage.getTableCellRange(firstRowString, headersRowCount)
                        .addRowsToTop(1),
                headerDescription,
                headersRowCount);
    }
}
