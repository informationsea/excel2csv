/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author yasu
 */
public abstract class ExcelTableWriterBase  implements TableWriter{
    
    private Workbook m_workbook = null;
    private Sheet m_sheet = null;
    private String m_path = null;
    private int m_rownum = 0;
    private String m_sheetname_candidate = "Sheet";
    private boolean m_overwrite_sheet = false;

    @Override
    public void close() throws IOException {
        try (FileOutputStream fos = new FileOutputStream(m_path)) {
            m_workbook.write(fos);
        }
    }
    
    public void setSheetNameCandidate(String sheetname, boolean overwriteSheet) {
        if (m_workbook != null)
            throw new IllegalStateException("Call setSheetNameCandidate before open");
        m_sheetname_candidate = sheetname;
        m_overwrite_sheet = overwriteSheet;
    }

    protected void open(String path, Workbook workbook) throws IOException {
        m_path = path;
        m_workbook = workbook;
        
        if (m_overwrite_sheet) {
            int num = workbook.getNumberOfSheets();
            for (int i = 0; i < num; ++i) {
                if (workbook.getSheetName(i).equals(m_sheetname_candidate)) {
                    workbook.removeSheetAt(i);
                    break;
                }
            }
        }

        boolean valid_name;
        int count = 1;
        do {
            try {
                if (count > 1) {
                    m_sheet = m_workbook.createSheet(m_sheetname_candidate + count);
                } else {
                    m_sheet = m_workbook.createSheet(m_sheetname_candidate);
                }
                valid_name = true;
            } catch (IllegalArgumentException iae) {
                valid_name = false;
            }
            count += 1;
            if (count > 20) {
                throw new IllegalArgumentException("Cannot make sheet "+m_sheetname_candidate);
            }
            
        } while (!valid_name);
        
        m_rownum = 0;
    }

    @Override
    public void writeRow(String[] row) throws IOException {
        Row r = m_sheet.createRow(m_rownum);
        for (int i = 0; i < row.length; i++) {
            Cell c = r.createCell(i);
            try {
                double v = Double.parseDouble(row[i]);
                c.setCellValue(v);
            } catch (NumberFormatException nfe) {
                c.setCellValue(row[i]);
            }
        }
        m_rownum += 1;
    }

}
