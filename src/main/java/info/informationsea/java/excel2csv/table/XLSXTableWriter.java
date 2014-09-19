/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author yasu
 */
public class XLSXTableWriter extends ExcelTableWriterBase{
    
    @Override
    public void open(String path) throws IOException {
        File file = new File(path);
        Workbook workbook;
        if (file.isFile()) {
            try (FileInputStream fis = new FileInputStream(path)) {
                workbook = new XSSFWorkbook(fis);
            }
        } else {
            workbook = new XSSFWorkbook();
        }
        open(path, workbook);
    }
}
