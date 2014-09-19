/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author yasu
 */
public class XLSXTableReader extends ExcelTableReaderBase{

   @Override
    public void open(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            Workbook workbook = new XSSFWorkbook(fis);
            open(workbook);
        }
    }
}
