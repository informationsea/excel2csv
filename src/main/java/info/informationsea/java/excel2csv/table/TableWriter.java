/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package info.informationsea.java.excel2csv.table;

import java.io.IOException;

/**
 *
 * @author yasu
 */
public interface TableWriter {

    /**
     *
     * @param path if null open standard output
     * @throws IOException
     */
    public abstract void open(String path) throws IOException;
    public abstract void writeRow(String[] row) throws IOException;
    public abstract void close() throws IOException;
}
