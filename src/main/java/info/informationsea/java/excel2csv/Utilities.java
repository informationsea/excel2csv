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
import info.informationsea.tableio.TableRecord;
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
import java.nio.file.FileVisitResult;
import java.nio.file.FileVisitor;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;

/**
 *
 * @author yasu
 */
@Slf4j
public class Utilities {
    public enum FileType {FILETYPE_XLSX, FILETYPE_XLS, FILETYPE_CSV, FILETYPE_TAB, FILETYPE_UNKNOWN};
    
    public static FileType suggestFileTypeFromName(String name) {
        if (name.endsWith(".xls"))
            return FileType.FILETYPE_XLS;
        if (name.endsWith(".xlsx"))
            return FileType.FILETYPE_XLSX;
        if (name.endsWith(".csv"))
            return FileType.FILETYPE_CSV;
        if (name.endsWith(".txt") || name.endsWith(".tsv"))
            return FileType.FILETYPE_TAB;
        return FileType.FILETYPE_UNKNOWN;
    }

    public static TableReader openReader(File inputFile, int sheetIndex, String sheetName) throws IOException {
        if (inputFile == null) {
            return new TableCSVReader(new InputStreamReader(System.in), new TabDelimitedFormat());
        } else {
            FileType type = suggestFileTypeFromName(inputFile.getName());
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

    public static TableWriter openWriter(final File outputFile, String sheetName, boolean overWrite, boolean disablePretty) throws IOException {
        if (outputFile == null) {
            return new TableCSVWriter(new OutputStreamWriter(System.out), new TabDelimitedFormat());
        } else {
            FileType type = suggestFileTypeFromName(outputFile.getName());
            switch (type) {
                case FILETYPE_XLS:
                case FILETYPE_XLSX: {
                    final Workbook workbook;
                    if (outputFile.exists() && outputFile.length() > 512) {
                        if (type == FileType.FILETYPE_XLSX)
                            workbook = new XSSFWorkbook(new FileInputStream(outputFile));
                        else workbook = new HSSFWorkbook(new FileInputStream(outputFile));
                    } else {
                        if (type == FileType.FILETYPE_XLSX)
                            workbook = new XSSFWorkbook();
                        else workbook = new HSSFWorkbook();
                    }

                    Sheet sheet = createUniqueNameSheetForWorkbook(workbook, sheetName, overWrite);
                    final ExcelSheetWriter excelSheetWriter = new ExcelSheetWriter(sheet);
                    excelSheetWriter.setPrettyTable(!disablePretty);
                    return new AbstractTableWriter() {
                        @Override
                        public void printRecord(Object... values) throws Exception {
                            for (int i = 0; i < values.length; i++) {
                                if (values[i] == null)
                                    values[i] = "";
                            }
                            excelSheetWriter.printRecord(values);
                        }

                        @Override
                        public void close() throws Exception {
                            excelSheetWriter.close();
                            try (OutputStream os = new FileOutputStream(outputFile)) {
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

    public static Sheet createUniqueNameSheetForWorkbook(Workbook workbook, String sheetName, boolean overwrite) {
        if (overwrite) {
            int index = workbook.getSheetIndex(workbook.getSheet(sheetName));
            if (index >= 0) workbook.removeSheetAt(index);
            return workbook.createSheet(sheetName);
        }

        String realSheetName = sheetName;
        int index = 1;
        Sheet sheet;
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
        return sheet;
    }

    public static void copyTable(TableReader reader, TableWriter writer) throws Exception {
        for (TableRecord record : reader) {
            writer.printRecord(record.getContent());
        }
    }

    public static void deleteDirectoryRecursive(Path dir) throws IOException {
        Files.walkFileTree(dir, new FileVisitor<Path>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                log.info("Delete {}", file);
                Files.delete(file);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) throws IOException {
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                log.info("Delete {}", dir);
                Files.delete(dir);
                return FileVisitResult.CONTINUE;
            }
        });
    }
}
