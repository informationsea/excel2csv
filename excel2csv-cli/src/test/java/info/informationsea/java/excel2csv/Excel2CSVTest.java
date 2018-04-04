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

import info.informationsea.tableio.TableCell;
import info.informationsea.tableio.TableReader;
import info.informationsea.tableio.csv.TableCSVReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
public class Excel2CSVTest {

    private static Path sampleDirectory;
    private static Path[] sampleInput;
    private static final String[] suffixList = new String[]{
            "txt", "csv", "xls", "xlsx"
    };
    private static List<TableCell[]> expectedData;

    @BeforeClass
    public static void setupClass() throws Exception {
        sampleInput = new Path[4];
        sampleDirectory = Files.createTempDirectory("excel2csv");

        int i = 0;
        for (String suffix : suffixList) {
            sampleInput[i] = Paths.get(sampleDirectory.toString(), "test."+suffix);
            try (InputStream is = Excel2CSVTest.class.getResourceAsStream("iris."+suffix)) {
                copyFile(is, sampleInput[i]);
            }
            i++;
        }

        expectedData = new TableCSVReader(new InputStreamReader(Excel2CSVTest.class.getResourceAsStream("iris.csv"))).readAll();
    }

    private static void copyFile(InputStream inputStream, Path output) throws IOException {
        try (OutputStream outputStream = new FileOutputStream(output.toFile())) {
            int readLength;
            byte[] buf = new byte[1024 * 10];

            while ((readLength = inputStream.read(buf)) > 0) {
                outputStream.write(buf, 0, readLength);
            }
        }
    }

    @AfterClass
    public static void teardownClass() throws Exception {
        Utilities.deleteDirectoryRecursive(sampleDirectory);
    }

    @Test
    public void testCopyFile() throws Exception {
        try (InputStream is = new FileInputStream(sampleInput[0].toFile())) {
            Assert.assertArrayEquals(IOUtils.toByteArray(Excel2CSVTest.class.getResourceAsStream("iris.txt")),
                    IOUtils.toByteArray(is));
        }
    }

    @Test
    public void testRun() throws Exception {
        for (Path input : sampleInput) {
            for (String suffix2 : suffixList) {
                Path outputFile = Files.createTempFile("outputTest", "."+suffix2);
                log.info("input: {}  / output: {}", input.getFileName().toString(), outputFile.getFileName().toString());
                new Excel2CSV().run(new String[]{input.toString(), outputFile.toString()});
                try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), 0, null)) {
                    assertObjects(expectedData, tableReader.readAll());
                }
            }
        }
    }

    @Test
    public void testMultiSheet() throws Exception {
        Path outputFile = Files.createTempFile("outputTest", ".xlsx");
        new Excel2CSV().run(new String[]{sampleInput[0].toString(), outputFile.toString()});
        new Excel2CSV().run(new String[]{sampleInput[1].toString(), outputFile.toString()});
        new Excel2CSV().run(new String[]{sampleInput[0].toString(), outputFile.toString()});

        try (FileInputStream inputStream = new FileInputStream(outputFile.toFile())) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                log.info("sheet {}", workbook.getSheetName(i));
            }
        }

        for (int i = 0; i < 3; i++) {
            try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), i, null)) {
                assertObjects(expectedData, tableReader.readAll());
            }
        }

        try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), 10, sampleInput[0].getFileName().toString())) {
            assertObjects(expectedData, tableReader.readAll());
        }

        try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), 10, sampleInput[1].getFileName().toString())) {
            assertObjects(expectedData, tableReader.readAll());
        }

        try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), 10, sampleInput[0].getFileName().toString()+"-1")) {
            assertObjects(expectedData, tableReader.readAll());
        }
    }

    @Test
    public void testNamedSheet() throws Exception {
        Path outputFile = Files.createTempFile("outputTest", ".xlsx");
        new Excel2CSV().run(new String[]{"-S", "OK", sampleInput[0].toString(), outputFile.toString()});
        try (TableReader tableReader = Utilities.openReader(outputFile.toFile(), 100, "OK")) {
            assertObjects(expectedData, tableReader.readAll());
        }
    }

    public static void assertObjects(List<TableCell[]> obj1, List<TableCell[]> obj2) {
        Assert.assertEquals(obj1.size(), obj2.size());

        for (int i = 0; i < obj1.size(); i++) {
            TableCell[] array1 = obj1.get(i);
            TableCell[] array2 = obj2.get(i);
            Assert.assertEquals(array1.length, array2.length);
            for (int j = 0; j < array1.length; j++) {
                String str1 = array1[j].toString();
                String str2 = array2[j].toString();

                try {
                    double num1 = Double.parseDouble(str1);
                    double num2 = Double.parseDouble(str2);
                    Assert.assertEquals(num1, num2, 0.00000001);
                } catch (NumberFormatException e) {
                    Assert.assertEquals(array1[j].toString(), array2[j].toString());
                }
            }
        }
    }
}