package info.informationsea.java.excel2csv;

import info.informationsea.tableio.TableWriter;
import info.informationsea.tableio.impl.AbstractTableWriter;

/**
 * excel2csv
 * Copyright (C) 2015 OKAMURA Yasunobu
 * Created on 15/03/04.
 */
public class FilteredWriter extends AbstractTableWriter {

    private TableWriter writer;

    public FilteredWriter(TableWriter writer) {
        this.writer = writer;
    }

    @Override
    public void printRecord(Object... values) throws Exception {
        Object[] filteredValues = new Object[values.length];
        for (int i = 0; i < values.length; i++) {
            try {
                filteredValues[i] = Double.valueOf(values[i].toString());
            } catch (NumberFormatException e) {
                filteredValues[i] = values[i];
            }
        }
        writer.printRecord(filteredValues);
    }

    @Override
    public void close() throws Exception {
        writer.close();
    }
}
