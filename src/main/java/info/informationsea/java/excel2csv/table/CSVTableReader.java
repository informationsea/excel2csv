/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import au.com.bytecode.opencsv.CSVReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;

/**
 *
 * @author yasu
 */
public class CSVTableReader implements TableReader {
    
    private CSVReader m_reader = null;

    @Override
    public String[] readRow() throws IOException {
        return m_reader.readNext();
    }

    @Override
    public void open(String path) throws IOException {
        if (path == null) {
            m_reader = new CSVReader(new InputStreamReader(System.in));
        } else {
            m_reader = new CSVReader(new FileReader(path));
        }
    }

    @Override
    public void close() throws IOException {
        m_reader.close();
    }
    
}
