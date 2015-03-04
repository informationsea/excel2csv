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

import info.informationsea.tableio.TableReader;
import info.informationsea.tableio.TableWriter;
import info.informationsea.tableio.csv.TableCSVReader;
import info.informationsea.tableio.csv.TableCSVWriter;
import info.informationsea.tableio.csv.format.DefaultFormat;
import info.informationsea.tableio.csv.format.TabDelimitedFormat;
import info.informationsea.tableio.excel.ExcelSheetReader;
import info.informationsea.tableio.excel.ExcelSheetWriter;
import info.informationsea.tableio.impl.AbstractTableWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 *
 * @author yasu
 */
@Slf4j
public class Utilities {
    public enum FileType {FILETYPE_XLSX, FILETYPE_XLS, FILETYPE_CSV, FILETYPE_TAB};
    
    public static FileType suggestFileTypeFromName(String name) {
        if (name.endsWith(".xls")) {
            return FileType.FILETYPE_XLS;
        } else if (name.endsWith(".xlsx")) {
            return FileType.FILETYPE_XLSX;
        } else if (name.endsWith(".csv")) {
            return FileType.FILETYPE_CSV;
        } else {
            return FileType.FILETYPE_TAB;
        }
    }

    public static TableReader openReader(String inputFile, int sheetIndex, String sheetName) throws IOException {
        if (inputFile == null) {
            return new TableCSVReader(new InputStreamReader(System.in), new TabDelimitedFormat());
        } else {
            FileType type = suggestFileTypeFromName(inputFile);
            switch (type) {
                case FILETYPE_XLS:
                case FILETYPE_XLSX: {
                    Workbook workbook;
                    if (type == FileType.FILETYPE_XLSX) workbook = new XSSFWorkbook(new FileInputStream(inputFile));
                    else workbook = new HSSFWorkbook(new FileInputStream(inputFile));

                    if (sheetName != null)
                        return new ExcelSheetReader(workbook.getSheet(sheetName));
                    else
                        return new ExcelSheetReader(workbook.getSheetAt(sheetIndex));
                }
                case FILETYPE_CSV:
                    return new TableCSVReader(new FileReader(inputFile), new DefaultFormat());
                case FILETYPE_TAB:
                default:
                    return new TableCSVReader(new FileReader(inputFile), new TabDelimitedFormat());
            }
        }
    }

    public static TableWriter openWriter(String outputFile, String sheetName, boolean overWrite, boolean disablePretty) throws IOException {
        if (outputFile == null) {
            return new TableCSVWriter(new OutputStreamWriter(System.out), new TabDelimitedFormat());
        } else {
            FileType type = suggestFileTypeFromName(outputFile);
            switch (type) {
                case FILETYPE_XLS:
                case FILETYPE_XLSX: {
                    final Workbook workbook;
                    final File outputFileObj = new File(outputFile);
                    if (outputFileObj.exists() && outputFileObj.length() > 512) {
                        if (type == FileType.FILETYPE_XLSX)
                            workbook = new XSSFWorkbook(new FileInputStream(outputFile));
                        else workbook = new HSSFWorkbook(new FileInputStream(outputFile));
                    } else {
                        if (type == FileType.FILETYPE_XLSX)
                            workbook = new XSSFWorkbook();
                        else workbook = new HSSFWorkbook();
                    }

                    Sheet sheet;
                    if (overWrite) {
                        int sheetIndex = workbook.getSheetIndex(sheetName);
                        if (sheetIndex >= 0)
                            workbook.removeSheetAt(sheetIndex);
                        sheet = workbook.createSheet(sheetName);
                    } else {
                        String realSheetName = sheetName;
                        int index = 1;
                        while (true) {
                            try {
                                sheet = workbook.createSheet(realSheetName);
                                break;
                            } catch (IllegalArgumentException e) {
                                realSheetName = sheetName + "-" + index++;
                                if (index > 20) {
                                    throw e;
                                }
                            }
                        }
                        log.info("new name {}", realSheetName);

                    }
                    final ExcelSheetWriter excelSheetWriter = new ExcelSheetWriter(sheet);
                    excelSheetWriter.setPrettyTable(!disablePretty);
                    return new AbstractTableWriter() {
                        @Override
                        public void printRecord(Object... values) throws Exception {
                            excelSheetWriter.printRecord(values);
                        }

                        @Override
                        public void close() throws Exception {
                            excelSheetWriter.close();
                            try (OutputStream os = new FileOutputStream(outputFileObj)) {
                                workbook.write(os);
                            }
                        }
                    };
                }
                case FILETYPE_CSV:
                    return new TableCSVWriter(new FileWriter(outputFile));
                case FILETYPE_TAB:
                default:
                    return new TableCSVWriter(new FileWriter(outputFile), new TabDelimitedFormat());
            }
        }
    }
}
