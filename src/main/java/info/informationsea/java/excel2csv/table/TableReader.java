/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.IOException;
import java.io.Reader;

/**
 *
 * @author yasu
 */
public interface TableReader {

    /**
     *
     * @param path if null, open standard input
     * @throws IOException
     */
    public abstract void open(String path) throws IOException;
    public abstract String[] readRow() throws IOException;
    public abstract void close() throws IOException;
}
