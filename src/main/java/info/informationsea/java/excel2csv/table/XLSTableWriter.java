/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author yasu
 */
public class XLSTableWriter extends ExcelTableWriterBase{
    
    @Override
    public void open(String path) throws IOException {
        File file = new File(path);
        Workbook workbook;
        if (file.isFile()) {
            try (FileInputStream fis = new FileInputStream(path)) {
                workbook = new HSSFWorkbook(fis);
            }
        } else {
            workbook = new HSSFWorkbook();
        }
        open(path, workbook);
    }
}
