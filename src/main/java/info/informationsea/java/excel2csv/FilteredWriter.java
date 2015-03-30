/*
 *  excel2csv  xls/xlsx/csv/tsv converter
 *  Copyright (C) 2015 Yasunobu OKAMURA
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

package info.informationsea.java.excel2csv;

import info.informationsea.tableio.TableWriter;
import info.informationsea.tableio.impl.AbstractTableWriter;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class FilteredWriter extends AbstractTableWriter {

    private TableWriter writer;
    private boolean autoConvert;

    @Override
    public void printRecord(Object... values) throws Exception {
        Object[] filteredValues = new Object[values.length];
        for (int i = 0; i < values.length; i++) {
            if (autoConvert) {
                try {
                    filteredValues[i] = Double.valueOf(values[i].toString());
                } catch (NumberFormatException e) {
                    filteredValues[i] = values[i];
                }
            } else {
                if (values[i] != null)
                    filteredValues[i] = values[i].toString();
                else
                    filteredValues[i] = "";
            }
        }
        writer.printRecord(filteredValues);
    }

    @Override
    public void close() throws Exception {
        writer.close();
    }
}
