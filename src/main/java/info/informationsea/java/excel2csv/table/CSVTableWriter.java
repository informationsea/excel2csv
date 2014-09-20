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

import au.com.bytecode.opencsv.CSVWriter;

import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;

/**
 *
 * @author yasu
 */
public class CSVTableWriter implements TableWriter{
    
    private CSVWriter m_writer = null;
	private char m_delimiter = ',';
	private char m_quoteChar = '"';
	
	public CSVTableWriter() {
		
	}

	public CSVTableWriter(char delimiter, char quoteChar) {
		m_delimiter = delimiter;
		m_quoteChar = quoteChar;
	}

    @Override
    public void open(String path) throws IOException {
        if (path == null) {
            open(new OutputStreamWriter(System.out));
        } else {
            open(new FileWriter(path));
        }
    }
    
    public void open(Writer writer) throws IOException {
    	m_writer = new CSVWriter(writer, m_delimiter, m_quoteChar);
    }

    @Override
    public void writeRow(String[] row) throws IOException {
        m_writer.writeNext(row);
    }

    @Override
    public void close() throws IOException {
        m_writer.close();
    }
    
}
