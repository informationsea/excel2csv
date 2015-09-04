/*
 *  excel2csv  xls/xlsx/csv/tsv converter
 *  Copyright (C) 2015 Yasunobu OKAMURA
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

package info.informationsea.java.excel2csv;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

public class UtilitiesTest {

    @Test
    public void testSuggestFileTypeFromName() throws Exception {
        Assert.assertEquals(Utilities.FileType.FILETYPE_CSV, Utilities.suggestFileTypeFromName("aa.csv"));
        Assert.assertEquals(Utilities.FileType.FILETYPE_XLS, Utilities.suggestFileTypeFromName("aa.xls"));
        Assert.assertEquals(Utilities.FileType.FILETYPE_XLSX, Utilities.suggestFileTypeFromName("aa.xlsx"));
        Assert.assertEquals(Utilities.FileType.FILETYPE_TAB, Utilities.suggestFileTypeFromName("aa.txt"));
        Assert.assertEquals(Utilities.FileType.FILETYPE_UNKNOWN, Utilities.suggestFileTypeFromName("aa.zip"));
    }

    @Test
    public void testCreateUniqueNameSheetForWorkbook() throws Exception {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet1 = Utilities.createUniqueNameSheetForWorkbook(workbook, "Sheet", false);
        Assert.assertEquals("Sheet", sheet1.getSheetName());
        Sheet sheet2 = Utilities.createUniqueNameSheetForWorkbook(workbook, "Sheet", false);
        Assert.assertEquals("Sheet-1", sheet2.getSheetName());
        Sheet sheet3 = Utilities.createUniqueNameSheetForWorkbook(workbook, "Sheet", false);
        Assert.assertEquals("Sheet-2", sheet3.getSheetName());
        Sheet sheet4 = Utilities.createUniqueNameSheetForWorkbook(workbook, "Sheet", true);
        Assert.assertEquals("Sheet", sheet4.getSheetName());
    }

    @Test
    public void testCopyTable() throws Exception {

    }
}