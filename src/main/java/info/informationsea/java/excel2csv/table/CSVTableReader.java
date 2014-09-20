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

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;

import au.com.bytecode.opencsv.CSVReader;

public class CSVTableReader implements TableReader {

	private char m_delimiter = ',';
	private char m_quoteChar = '"';
    private CSVReader m_reader = null;
    private String m_path = null;
    
    public CSVTableReader() {
		
	}
    
    public CSVTableReader(char delimiter, char quoteChar) {
		m_delimiter = delimiter;
		m_quoteChar = quoteChar;
	}
    
    @Override
    public void open(String path) throws IOException {
        if (path == null) {
            open(new InputStreamReader(System.in));
        } else {
            open(new FileReader(path));
            m_path = path;
        }
    }
    
    public void open(Reader reader) {
    	m_reader = new CSVReader(reader, m_delimiter, m_quoteChar);
    }

    @Override
    public String[] readRow() throws IOException {
        return m_reader.readNext();
    }
    
    @Override
    public void close() throws IOException {
        m_reader.close();
    }

	@Override
	public void setSheetName(String sheetName) throws IllegalArgumentException {
		// ignore
	}

	@Override
	public void setSheetIndex(int sheetIndex) throws IllegalArgumentException {
		// ignore
	}

	@Override
	public String[] getSheetList() {
		if (m_path == null)
			return new String[]{"Sheet"};
		return new String[]{new File(m_path).getName()};
	}
}
