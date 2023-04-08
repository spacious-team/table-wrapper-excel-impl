/*
 * Table Wrapper Excel Impl
 * Copyright (C) 2023  Spacious Team <spacious-team@ya.ru>
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
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.math.BigDecimal;
import java.time.Instant;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class ExcelCellDataAccessObjectTest {

    @Mock
    Cell cell;
    ExcelCellDataAccessObject dao = spy(ExcelCellDataAccessObject.INSTANCE);

    @Test
    void getCell() {
        Row row = mock(Row.class);
        ExcelTableRow tableRow = ExcelTableRow.of(row);

        dao.getCell(tableRow, 1);

        verify(row).getCell(1);
    }

    @Test
    void getBigDecimalValue() {
        doReturn(1.1).when(dao).getDoubleValue(cell);

        BigDecimal actual = dao.getBigDecimalValue(cell);

        assertEquals(BigDecimal.valueOf(1.1), actual);
        verify(dao).getDoubleValue(cell);
    }

    @Test
    void getStringValue() {
        doReturn("test").when(dao).getValue(cell);
        assertEquals("test", dao.getStringValue(cell));
    }

    @Test
    void getStringValue_numericValue() {
        doReturn(1.0).when(dao).getValue(cell);
        assertEquals("1", dao.getStringValue(cell));
    }

    @Test
    void getInstantValue() {
        Date date = new Date();
        Instant expected = date.toInstant();
        when(cell.getDateCellValue()).thenReturn(date);

        assertEquals(expected, dao.getInstantValue(cell));

        verify(cell).getDateCellValue();
    }

    @Test
    void getLocalDateTimeValue() {
        dao.getLocalDateTimeValue(cell);
        verify(cell).getLocalDateTimeCellValue();
    }
}