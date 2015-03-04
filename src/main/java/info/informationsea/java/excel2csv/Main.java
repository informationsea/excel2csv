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

import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import info.informationsea.tableio.TableReader;
import info.informationsea.tableio.TableRecord;
import info.informationsea.tableio.TableWriter;
import lombok.extern.slf4j.Slf4j;
import org.kohsuke.args4j.Argument;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;


@Slf4j
public class Main {

    @Option(name = "-s", metaVar = "SHEET", usage = "Sheet name of input file (xls/xlsx only)")
    String optionSheetName = null;

    @Option(name = "-i", metaVar = "INDEX", usage = "Sheet index of input file (xls/xlsx only / default: 0 / start from 0)")
    int optionSheetIndex = 0;

    @Option(name = "-S", metaVar = "SHEET", usage = "Sheet name candidate of output file (xls/xlsx only)")
    String optionOutputSheetName = null;

    @Option(name = "-F", usage = "Overwrite sheet if exists")
    boolean optionOverwrite = false;

    @Option(name = "-h", usage = "Show help")
    boolean optionHelp = false;

    @Option(name = "-p", usage = "Disable pretty table (xls/xlsx only)")
    boolean optionDisablePretty = false;

    @Argument
    private List<String> arguments = new ArrayList<String>();

    public static void main(String[] argv) {
        new Main().run(argv);
    }


    public void run(String[] argv) {
        CmdLineParser cmdLineParser = new CmdLineParser(this);
        try {
            cmdLineParser.parseArgument(argv);
        } catch (CmdLineException e) {
            System.err.println(e.getLocalizedMessage());
            optionHelp = true;
        }

        if (optionHelp) {
            System.err.println("excel2csv [options] [INPUT [OUTPUT]]");
            cmdLineParser.printUsage(System.err);
            return;
        }

        String inputFile = arguments.size() >= 1 ? arguments.get(0) : null;
        String outputFile = arguments.size() >= 2 ? arguments.get(1) : null;

        if (optionOutputSheetName == null) {
            if (inputFile != null && !inputFile.equals("-")) {
                optionOutputSheetName = new File(inputFile).getName();
            } else {
                optionOutputSheetName = "Sheet";
            }
        }
        
        // END OF PARSING OPTIONS
        
        try (TableReader reader = Utilities.openReader(inputFile, optionSheetIndex, optionSheetName)) {
            try (TableWriter writer = new FilteredWriter(Utilities.openWriter(outputFile, optionOutputSheetName, optionOverwrite, optionDisablePretty))) {
                for (TableRecord record : reader) {
                    writer.printRecord(record.getContent());
                }
            }
        } catch (Exception ex) {
            log.error("Error on writing {}", ex);
        }
    }

}
