/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author yasu
 */
public class XLSTableReader extends ExcelTableReaderBase{

   @Override
    public void open(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            Workbook workbook = new HSSFWorkbook(fis);
            open(workbook);
        }
    }
}
