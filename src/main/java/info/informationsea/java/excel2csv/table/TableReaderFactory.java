/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
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
