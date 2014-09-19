/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author yasu
 */
public abstract class ExcelTableReaderBase implements TableReader {
    
    private Workbook m_workbook;
    private Sheet m_sheet;
    private Iterator<Row> m_rowIterator;
    
    private int m_sheetIndex = -1;
    private String m_sheetName = null;
    
    public void setSheetName(String sheetName) {
        if (m_workbook != null)
            throw new IllegalStateException("Call setSheetName before open");
        m_sheetName = sheetName;
    }
    
    public void setSheetIndex(int sheetIndex) {
        if (m_workbook != null)
            throw new IllegalStateException("Call setSheetIndex before open");
        m_sheetIndex = sheetIndex;
    }
    
    public void open(Workbook workbook) throws IOException {
        m_workbook = workbook;
        if (m_sheetName != null)
            m_sheet = m_workbook.getSheet(m_sheetName);
        else if (m_sheetIndex >= 0)
            m_sheet = m_workbook.getSheetAt(m_sheetIndex);
        else
            m_sheet = m_workbook.getSheetAt(0);
        
        m_rowIterator = m_sheet.rowIterator();
    }

    @Override
    public String[] readRow() throws IOException {
        if (!m_rowIterator.hasNext()) {
            return null;
        }
        
        ArrayList<String> rowList = new ArrayList<>();
        Row row = m_rowIterator.next();
        
        for (Iterator<Cell> cell_iterator = row.cellIterator(); cell_iterator.hasNext();) {
            Cell cell = cell_iterator.next();
            while (rowList.size() < cell.getColumnIndex()) {
                rowList.add("");
            }
            
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    rowList.add(Boolean.toString(cell.getBooleanCellValue()));
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    rowList.add(Double.toString(cell.getNumericCellValue()));
                    break;
                case Cell.CELL_TYPE_STRING:
                    rowList.add(cell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    rowList.add(cell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_BLANK:
                default:
                    rowList.add("");
                    break;
            }
        }
        
        String[] rowArray = new String[rowList.size()];
        rowList.toArray(rowArray);
        return rowArray;
    }

    @Override
    public void close() throws IOException {
        // nothing to do
    }
    
}
