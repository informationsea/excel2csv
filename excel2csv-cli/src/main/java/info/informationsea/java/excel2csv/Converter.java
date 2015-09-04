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

import info.informationsea.tableio.TableReader;
import info.informationsea.tableio.TableWriter;
import info.informationsea.tableio.excel.ExcelSheetReader;
import info.informationsea.tableio.excel.ExcelSheetWriter;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;


@Builder @Data @NoArgsConstructor @AllArgsConstructor @Slf4j
public class Converter {
    private boolean overwriteSheet = false;
    private boolean copyAllSheets = true;
    private String inputSheetName = null;
    private int inputSheetIndex = 0;
    private String outputSheetName = "Sheet";

    private boolean prettyTable = true;
    private boolean convertCellTypes = true;

    public void doConvert(List<File> inputFiles, File outputFile) throws Exception {
        if (copyAllSheets || inputFiles.size() > 1) {
            switch (Utilities.suggestFileTypeFromName(outputFile.getName())) {
                case FILETYPE_XLS:
                case FILETYPE_XLSX:
                    doConvertAllSheets(inputFiles, outputFile);
                    break;
                default:
                    throw new IllegalArgumentException("Output file format should be Excel format");
            }
        } else {
            doConvertOne(inputFiles.get(0), outputFile);
        }
    }

    private void doConvertAllSheets(List<File> inputFiles, File outputFile) throws Exception {
        Workbook workbook;

        if (outputFile.isFile() && outputFile.length() > 512) {
            switch (Utilities.suggestFileTypeFromName(outputFile.getName())) {
                case FILETYPE_XLS:
                case FILETYPE_XLSX:
                    workbook = WorkbookFactory.create(outputFile);
                    break;
                default:
                    throw new IllegalArgumentException("Output file format should be Excel format");
            }
        } else {
            switch (Utilities.suggestFileTypeFromName(outputFile.getName())) {
                case FILETYPE_XLS:
                    workbook = new HSSFWorkbook();
                    break;
                case FILETYPE_XLSX:
                    workbook = new XSSFWorkbook();
                    break;
                default:
                    throw new IllegalArgumentException("Output file format should be Excel format");
            }
        }

        for (File oneInput : inputFiles) {
            switch (Utilities.suggestFileTypeFromName(oneInput.getName())) {
                case FILETYPE_XLSX:
                case FILETYPE_XLS: {
                    Workbook inputWorkbook = WorkbookFactory.create(oneInput);
                    int sheetNum = inputWorkbook.getNumberOfSheets();
                    for (int i = 0; i < sheetNum; i++) {
                        try (TableReader reader = new ExcelSheetReader(inputWorkbook.getSheetAt(i))) {
                            ExcelSheetWriter sheetWriter = new ExcelSheetWriter(Utilities.createUniqueNameSheetForWorkbook(workbook, inputWorkbook.getSheetName(i), overwriteSheet));
                            sheetWriter.setPrettyTable(prettyTable);
                            try (TableWriter tableWriter = new FilteredWriter(sheetWriter, convertCellTypes)) {
                                Utilities.copyTable(reader, tableWriter);
                            }
                        }
                    }
                    break;
                }
                default: {
                    try (TableReader reader = Utilities.openReader(oneInput, inputSheetIndex, inputSheetName)) {
                        ExcelSheetWriter sheetWriter = new ExcelSheetWriter(Utilities.createUniqueNameSheetForWorkbook(workbook, oneInput.getName(), overwriteSheet));
                        sheetWriter.setPrettyTable(prettyTable);
                        try (TableWriter tableWriter = new FilteredWriter(sheetWriter, convertCellTypes)) {
                            Utilities.copyTable(reader, tableWriter);
                        }
                    }
                    break;
                }
            }
        }

        workbook.write(new FileOutputStream(outputFile));
    }

    private void doConvertOne(File inputFile, File outputFile) throws Exception {
        try (TableWriter writer = Utilities.openWriter(outputFile, outputSheetName, overwriteSheet, prettyTable)) {
            try (TableReader reader = Utilities.openReader(inputFile, inputSheetIndex, inputSheetName)) {
                try (FilteredWriter writer2 = new FilteredWriter(writer, convertCellTypes)) {
                    Utilities.copyTable(reader, writer2);
                }
            }
        }
    }
}
