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
public class TableWriterFactory {
    private TableWriterFactory(){}
    
    
    public static TableWriter openWriter(String path) throws IOException {
        return openWriter(path, "Sheet", false);
    }
    
    public static TableWriter openWriter(String path, String sheetNameCandidate, boolean overwriteSheet) throws IOException {
        TableWriter writer;
        switch (Utilities.suggestFileTypeFromName(path)) {
            case FILETYPE_CSV:
                writer = new CSVTableWriter();
                break;
            case FILETYPE_XLS:
                writer = new XLSTableWriter();
                ((XLSTableWriter)writer).setSheetNameCandidate(sheetNameCandidate, overwriteSheet);
                break;
            case FILETYPE_XLSX:
                writer = new XLSXTableWriter();
                ((XLSXTableWriter)writer).setSheetNameCandidate(sheetNameCandidate, overwriteSheet);
                break;
            default:
                writer = new TabTableWriter();
                break;
        }
        writer.open(path);
        return writer;
    }
}
