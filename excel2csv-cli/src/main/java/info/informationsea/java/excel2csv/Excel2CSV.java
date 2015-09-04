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

import lombok.extern.slf4j.Slf4j;
import org.kohsuke.args4j.Argument;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;

import java.io.File;
import java.util.ArrayList;
import java.util.List;


@Slf4j
public class Excel2CSV {

    @Option(name = "-a", usage = "Copy all sheets in input excel file")
    private boolean optionCopyAllSheets = false;

    @Option(name = "-s", metaVar = "SHEET", usage = "Sheet name of input file (xls/xlsx only)")
    private String optionSheetName = null;

    @Option(name = "-i", metaVar = "INDEX", usage = "Sheet index of input file (xls/xlsx only / default: 0 / start from 0)")
    private int optionSheetIndex = 0;

    @Option(name = "-S", metaVar = "SHEET", usage = "Sheet name candidate of output file (xls/xlsx only)")
    private String optionOutputSheetName = null;

    @Option(name = "-F", usage = "Overwrite sheet if exists")
    private boolean optionOverwrite = false;

    @Option(name = "-h", usage = "Show help")
    private boolean optionHelp = false;

    @Option(name = "-p", usage = "Disable pretty table (xls/xlsx output only)")
    private boolean optionDisablePretty = false;

    @Option(name = "-c", usage = "Disable cell types conversion automatically")
    private boolean optionDisableConvertCell = false;


    @Option(name = "-v", usage = "Show version")
    private boolean optionVersion = false;

    @Argument
    private List<String> arguments = new ArrayList<String>();

    public static void main(String[] argv) {
        new Excel2CSV().run(argv);
    }

    public void run(String[] argv) {
        CmdLineParser cmdLineParser = new CmdLineParser(this);
        try {
            cmdLineParser.parseArgument(argv);
        } catch (CmdLineException e) {
            System.err.println(e.getLocalizedMessage());
            optionHelp = true;
        }

        if (optionVersion) {
            System.err.println("Excel2CSV\nVersion: "+ VersionResolver.getVersion() +  "\n" +
                    "Git Commit: " + VersionResolver.getGitCommit() + "\n" +
                    "Build Date: " + VersionResolver.getBuildDate() + "\n\n" +
                    "Webpage: https://github.com/informationsea/excel2csv");
            return;
        }

        if (optionHelp || arguments.size() < 1) {
            System.err.println("excel2csv [options] INPUT... OUTPUT");
            cmdLineParser.printUsage(System.err);
            return;
        }

        if (optionOutputSheetName == null) {
            if (arguments.size() == 2 && !arguments.get(0).equals("-")) {
                optionOutputSheetName = new File(arguments.get(0)).getName();
            } else {
                optionOutputSheetName = "Sheet";
            }
        }
        
        // END OF PARSING OPTIONS
        List<File> inputFiles = new ArrayList<>();
        for (String one : arguments.subList(0, arguments.size()-1)) {
            inputFiles.add(new File(one));
        }
        File outputFile = new File(arguments.get(arguments.size()-1));
        try {
            Converter converter = Converter.builder().
                    convertCellTypes(!optionDisableConvertCell).
                    copyAllSheets(optionCopyAllSheets).
                    inputSheetIndex(optionSheetIndex).
                    inputSheetName(optionSheetName).
                    outputSheetName(optionOutputSheetName).
                    overwriteSheet(optionOverwrite).
                    prettyTable(!optionDisablePretty).build();
            //log.info("pretty {}", converter.isPrettyTable());
            converter.doConvert(inputFiles, outputFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
