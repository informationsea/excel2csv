/*
 *  excel2csv  xls/xlsx/csv/tsv converter
 *  Copyright (C) 2014 Yasunobu OKAMURA
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

package info.informationsea.java.excel2csv.table;

import java.io.IOException;

/**
 *
 * @author yasu
 */
public class TableReaderFactory {
    private TableReaderFactory(){};
    
    public static TableReader openReader(String path) throws IOException {
        return openReader(path, -1, null);
    }
    
    public static TableReader openReader(String path, int sheetIndex) throws IOException {
        return openReader(path, sheetIndex, null);
    }
    public static TableReader openReader(String path, String sheetName) throws IOException {
        return openReader(path, -1, sheetName);
    }
    
    private static TableReader openReader(String path, int sheetIndex, String sheetName) throws IOException {
        TableReader reader;
        switch (Utilities.suggestFileTypeFromName(path)) {
            case FILETYPE_CSV:
                reader = new CSVTableReader();
                break;
            case FILETYPE_XLS:
                reader = new XLSTableReader();
                break;
            case FILETYPE_XLSX:
                reader = new XLSXTableReader();
                break;
            default:
                reader = new TabTableReader();
                break;
        }
        
        try {
            if (sheetIndex >= 0)
                ((ExcelTableReaderBase)reader).setSheetIndex(sheetIndex);
            if (sheetName != null)
                ((ExcelTableReaderBase)reader).setSheetName(sheetName);
        } catch (ClassCastException cce) {} // ignore exception
        
        reader.open(path);
        return reader;
    }
    
    

}
