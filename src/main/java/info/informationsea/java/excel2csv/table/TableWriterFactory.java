/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.FileWriter;
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
