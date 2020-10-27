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

import org.spacious_team.table_wrapper.api.*;

public class ExcelTableFactory implements TableFactory {
    @Override
    public boolean canHandle(ReportPage reportPage) {
        return (reportPage instanceof ExcelSheet);
    }

    @Override
    public Table create(ReportPage reportPage, String tableName, String tableFooterString,
                        Class<? extends TableColumnDescription> headerDescription,
                        int headersRowCount) {
        AbstractTable table = new ExcelTable(reportPage, tableName,
                reportPage.getTableCellRange(tableName, headersRowCount, tableFooterString),
                headerDescription,
                headersRowCount);
        table.setLastTableRowContainsTotalData(true);
        return table;
    }

    @Override
    public Table create(ReportPage reportPage, String tableName,
                        Class<? extends TableColumnDescription> headerDescription,
                        int headersRowCount) {
        AbstractTable table = new ExcelTable(reportPage, tableName,
                reportPage.getTableCellRange(tableName, headersRowCount),
                headerDescription,
                headersRowCount);
        table.setLastTableRowContainsTotalData(false);
        return table;
    }

    @Override
    public Table createOfNoName(ReportPage reportPage, String madeUpTableName, String firstLineText,
                                Class<? extends TableColumnDescription> headerDescription,
                                int headersRowCount) {
        AbstractTable table = new ExcelTable(reportPage, madeUpTableName,
                getNoNameTableRange(reportPage, firstLineText, headersRowCount),
                headerDescription,
                headersRowCount);
        table.setLastTableRowContainsTotalData(true);
        return table;
    }
}
