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
import info.informationsea.tableio.TableCell;
import info.informationsea.tableio.csv.TableCSVReader;
import info.informationsea.tableio.csv.format.DefaultFormat;
import info.informationsea.tableio.excel.ExcelSheetReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

@Slf4j
public class ConverterTest {

    private static Path temporaryDirectory;

    private static Map<String, List<TableCell[]>> reference;

    @BeforeClass
    public static void beforeClass() throws Exception {
        temporaryDirectory = Files.createTempDirectory("excel2csv");
        for (String one : new String[]{"multisheet.xls", "multisheet.xlsx", "CO2.csv", "DNase.csv", "iris.csv"}) {
            try(InputStream is = ConverterTest.class.getResourceAsStream(one)) {
                try (OutputStream os = new FileOutputStream(new File(temporaryDirectory.toFile(), one))) {
                    IOUtils.copy(is, os);
                }
            }
        }

        reference = new HashMap<>();
        for (String one : new String[]{"iris", "CO2", "DNase"}) {
            try (TableCSVReader reader = new TableCSVReader(new InputStreamReader(ConverterTest.class.getResourceAsStream(one+".csv")), new DefaultFormat())) {
                reference.put(one, reader.readAll());
            }
        }
    }

    @AfterClass
    public static void afterClass() throws Exception {
        log.info("Delete Dir {}", temporaryDirectory);
        Utilities.deleteDirectoryRecursive(temporaryDirectory);
    }

    @Test
    public void testDoConvert() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(true).build().doConvert(Arrays.asList(
                new File(temporaryDirectory.toFile(), "CO2.csv"),
                new File(temporaryDirectory.toFile(), "DNase.csv"),
                new File(temporaryDirectory.toFile(), "iris.csv")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        for (String one : new String[]{"iris", "CO2", "DNase"}) {
            try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one + ".csv"))) {
                MainTest.assertObjects(reference.get(one), reader.readAll());
            }
        }
    }

    @Test
    public void testDoConvert2() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(true).build().doConvert(Arrays.asList(
                new File(temporaryDirectory.toFile(), "multisheet.xls"),
                new File(temporaryDirectory.toFile(), "multisheet.xlsx")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        for (String one : new String[]{"iris", "CO2", "DNase"}) {
            for (String suffix : new String[]{"", "-1"}) {
                try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one+suffix))) {
                    MainTest.assertObjects(reference.get(one), reader.readAll());
                }
            }
        }
    }

    @Test
    public void testDoConvert3() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(true).build().doConvert(Collections.singletonList(
                new File(temporaryDirectory.toFile(), "multisheet.xls")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        for (String one : new String[]{"iris", "CO2", "DNase"}) {
                try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one))) {
                    MainTest.assertObjects(reference.get(one), reader.readAll());
            }
        }
    }

    @Test
    public void testDoConvert4() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(false).outputSheetName("iris").build().doConvert(Collections.singletonList(
                new File(temporaryDirectory.toFile(), "multisheet.xls")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        Assert.assertEquals(1, workbook.getNumberOfSheets());
        for (String one : new String[]{"iris"}) {
            log.info("Sheet name: {}", workbook.getSheetName(0));
            try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one))) {
                MainTest.assertObjects(reference.get(one), reader.readAll());
            }
        }
    }

    @Test
    public void testDoConvert5() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(false).inputSheetName("DNase").outputSheetName("DNase").build().doConvert(Collections.singletonList(
                new File(temporaryDirectory.toFile(), "multisheet.xls")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        Assert.assertEquals(1, workbook.getNumberOfSheets());
        for (String one : new String[]{"DNase"}) {
            try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one))) {
                MainTest.assertObjects(reference.get(one), reader.readAll());
            }
        }
    }

    @Test
    public void testDoConvert6() throws Exception {
        Path temporaryOutput = Files.createTempFile("excel2csv", ".xlsx");
        Converter.builder().copyAllSheets(false).inputSheetIndex(2).outputSheetName("DNase").build().doConvert(Collections.singletonList(
                new File(temporaryDirectory.toFile(), "multisheet.xls")
        ), temporaryOutput.toFile());

        Workbook workbook = WorkbookFactory.create(temporaryOutput.toFile());
        Assert.assertEquals(1, workbook.getNumberOfSheets());
        for (String one : new String[]{"DNase"}) {
            try (TableReader reader = new ExcelSheetReader(workbook.getSheet(one))) {
                MainTest.assertObjects(reference.get(one), reader.readAll());
            }
        }
    }
}