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

package info.informationsea.java.excel2csv;

import info.informationsea.java.excel2csv.table.TabTableReader;
import info.informationsea.java.excel2csv.table.TabTableWriter;
import info.informationsea.java.excel2csv.table.TableReader;
import info.informationsea.java.excel2csv.table.TableReaderFactory;
import info.informationsea.java.excel2csv.table.TableWriter;
import info.informationsea.java.excel2csv.table.TableWriterFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.PosixParser;


/**
 *
 * @author yasu
 */
public class Main {
    public static void main(String[] argv) {
        Options options = new Options();
        options.addOption("a", false, "file type of input file will detect automatically (default)");
        options.addOption("A", false, "file type of output file will detect automatically (default)");
        options.addOption("s", true, "Sheet name of input file (xls/xlsx only)");
        options.addOption("i", true, "Sheet index of input file (xls/xlsx only / default: 0)");
        options.addOption("S", true, "Sheet name candidate of output file (xls/xlsx only)");
        options.addOption("F", false, "Overwrite sheet if exists");
        
        options.addOption("h", false, "show help");
        CommandLineParser parser = new PosixParser();
        CommandLine cmd;
        try {
            cmd = parser.parse(options, argv);
        } catch (ParseException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
            return;
        }
        
        if (cmd.hasOption("h")) {
            HelpFormatter formatter = new HelpFormatter();
            formatter.printHelp("excel2csv [options] [INPUT] [OUTPUT]", options);
            return;
        }
        
        String[] fileArgv = cmd.getArgs();
        String inputFile = fileArgv.length >= 1 ? fileArgv[0] : null;
        String outputFile = fileArgv.length >= 2 ? fileArgv[1] : null;
        
        String inputSheetName = cmd.getOptionValue("s", null);
        int inputSheetIndex = Integer.parseInt(cmd.getOptionValue("i", "-1"));
        String outputSheetName = cmd.getOptionValue("S", null);
        boolean overwriteSheet = cmd.hasOption("F");
        
        if (outputSheetName == null) {
            if (inputFile != null && !inputFile.equals("-")) {
                outputSheetName = new File(inputFile).getName();
            } else {
                outputSheetName = "Sheet";
            }
        }
        
        // END OF PARSING OPTIONS
        
        try {
            TableReader reader;
            TableWriter writer;
            if (inputFile != null && !inputFile.equals("-")) {
                if (inputSheetName != null)
                    reader = TableReaderFactory.openReader(inputFile, inputSheetName);
                else if (inputSheetIndex >= 0)
                    reader = TableReaderFactory.openReader(inputFile, inputSheetIndex);
                else
                    reader = TableReaderFactory.openReader(inputFile);
            } else {
                reader = new TabTableReader();
                ((TabTableReader)reader).open(new InputStreamReader(System.in));
            }
            
            if (outputFile != null && !outputFile.equals("-")) {
                writer = TableWriterFactory.openWriter(outputFile, outputSheetName, overwriteSheet);
            } else {
                writer = new TabTableWriter();
                ((TabTableWriter)writer).open(new OutputStreamWriter(System.out));
            }
            
            String[] row;
            while ((row = reader.readRow()) != null) {
                //System.err.printf("Write row: %s\n", row);
                writer.writeRow(row);
            }
            
            reader.close();
            writer.close();
        } catch (IOException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
