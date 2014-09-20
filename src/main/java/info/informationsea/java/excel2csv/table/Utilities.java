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

/**
 *
 * @author yasu
 */
public class Utilities {
    public enum FileType {FILETYPE_XLSX, FILETYPE_XLS, FILETYPE_CSV, FILETYPE_TAB};
    
    public static FileType suggestFileTypeFromName(String name) {
        if (name.endsWith(".xls")) {
            return FileType.FILETYPE_XLS;
        } else if (name.endsWith(".xlsx")) {
            return FileType.FILETYPE_XLSX;
        } else if (name.endsWith(".csv")) {
            return FileType.FILETYPE_CSV;
        } else {
            return FileType.FILETYPE_TAB;
        }
    }
}
