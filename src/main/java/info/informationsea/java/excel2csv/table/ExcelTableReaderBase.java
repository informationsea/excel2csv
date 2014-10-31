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
        m_sheetName = sheetName;
        if (m_workbook != null)
        	m_sheet = m_workbook.getSheet(sheetName);
    }
    
    public void setSheetIndex(int sheetIndex) {
    	m_sheetIndex = sheetIndex;
    	if (m_workbook != null)
    		m_sheet = m_workbook.getSheetAt(sheetIndex);
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
        
        ArrayList<String> rowList = new ArrayList<String>();
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

	@Override
	public String[] getSheetList() {
		String[] sheetList = new String[m_workbook.getNumberOfSheets()];
		for (int i = 0; i < sheetList.length; i++) {
			sheetList[i] = m_workbook.getSheetName(i);
		}
		return sheetList;
	}
    
}
