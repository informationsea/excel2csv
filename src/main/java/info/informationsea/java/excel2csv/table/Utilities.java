/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
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
