package net.cbean.office.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class XSSFSheetWriterTest {
    private XSSFSheetWriter sheetWriter;
    SXSSFWorkbook wb = null;
    FileOutputStream fos = null;
    String outputExcel = "target/test.xlsx";

    @Before
    public void setup() {
        wb = new SXSSFWorkbook();

        String[] headers = new String[]{"Name", "desc", "Num"};
        SXSSFSheet sh = wb.createSheet("sheet1");
        sheetWriter = new XSSFSheetWriter(sh, headers);
    }

    private void saveToFile() {
        try {
            fos = new FileOutputStream(outputExcel);
            wb.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            if (fos != null) {
                fos.close();
            }
        } catch (IOException e) {
        }
        try {
            if (wb != null) {
                wb.close();
            }
        } catch (IOException e) {
        }
    }

    private void assertOutputExcel(String outputExcel) throws Exception {
        File outFile = new File(outputExcel);
        FileInputStream in = new FileInputStream(outFile);
        Sheet sheet = new XSSFWorkbook(in).getSheetAt(0);
        assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("desc", sheet.getRow(0).getCell(1).getStringCellValue());
        assertEquals("Num", sheet.getRow(0).getCell(2).getStringCellValue());

        assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("b", sheet.getRow(1).getCell(1).getStringCellValue());
        assertTrue(sheet.getRow(1).getCell(2).getNumericCellValue() == 1d);

        assertEquals("a1", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals("b1", sheet.getRow(2).getCell(1).getStringCellValue());
        assertTrue(sheet.getRow(2).getCell(2).getNumericCellValue() == 2d);

        assertEquals("a2", sheet.getRow(3).getCell(0).getStringCellValue());
        assertEquals("b2", sheet.getRow(3).getCell(1).getStringCellValue());
        assertTrue(sheet.getRow(3).getCell(2).getNumericCellValue() == 3d);

        outFile.delete();
    }

    @Test
    public void testHeaderLabelNotMatch() throws Exception {
        String[] headers = new String[]{"Name", "desc", "Num"};
        SXSSFSheet sh = wb.createSheet("sheet2");
        try {
            new XSSFSheetWriter(sh, headers, new String[]{});
        } catch (IllegalArgumentException e) {
            assertEquals("Size of headerLabel is not match the size of header!", e.getMessage());
        }
    }

    @Test
    public void testAddDataWithValueHandler() throws Exception {
        sheetWriter.init();
        for (Object[] row : initData()) {
            sheetWriter.addRow((header, index) -> row[index]);
        }

        saveToFile();
        assertOutputExcel(outputExcel);
    }

    @Test
    public void testAddDataListObject() throws Exception {
        List<TestModel> data = new ArrayList<>();
        data.add(new TestModel(new String("a"), new String("b"), 1));
        data.add(new TestModel(new String("a1"), new String("b1"), 2));
        data.add(new TestModel(new String("a2"), new String("b2"), 3));

        sheetWriter.init().addData(data);

        saveToFile();
        assertOutputExcel(outputExcel);
    }

    @Test
    public void testAddDataList() throws Exception {
        sheetWriter.init().addData(initData());
        saveToFile();
        assertOutputExcel(outputExcel);
    }

    @Test
    public void testSetStyleWriter() throws Exception {
        sheetWriter.setStyleWriter("Name", cell -> {
            CellStyle style = wb.createCellStyle();

            // 设置前景色
            style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            // 设置颜色填充方式
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            setBorderStyle(style);
            cell.setCellStyle(style);
        }, cell -> {
            CellStyle style = wb.createCellStyle();
            setBorderStyle(style);
            cell.setCellStyle(style);
        }).init();
        sheetWriter.addData(initData());

        saveToFile();

        File outFile = new File(outputExcel);
        FileInputStream in = new FileInputStream(outFile);
        Sheet sheet = new XSSFWorkbook(in).getSheetAt(0);
        CellStyle headerStyle = sheet.getRow(0).getCell(0).getCellStyle();
        assertEquals(FillPatternType.SOLID_FOREGROUND, headerStyle.getFillPatternEnum());

        assertOutputExcel(outputExcel);
    }

    @Test
    public void testSetStyleWriterNegative() throws Exception {
        sheetWriter.setStyleWriter("Name", null, cell -> {
            CellStyle style = wb.createCellStyle();
            setBorderStyle(style);
            cell.setCellStyle(style);
        }).setStyleWriter("desc", cell -> {
            CellStyle style = wb.createCellStyle();
            setBorderStyle(style);
            cell.setCellStyle(style);
        }).init();
    }

    private void setBorderStyle(CellStyle style) {
        // 设置边框样式
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
    }

    @Test
    public void testAddDataArray() throws Exception {
        sheetWriter.init().addData(initData().toArray(new Object[]{}));

        saveToFile();
        assertOutputExcel(outputExcel);
    }

    private List<Object[]> initData() {
        List<Object[]> data = new ArrayList<>();
        data.add(new Object[]{new String("a"), new String("b"), 1});
        data.add(new Object[]{new String("a1"), new String("b1"), 2});
        data.add(new Object[]{new String("a2"), new String("b2"), 3});
        return data;
    }

    public class TestModel {
        String name;
        String Desc;
        int num;

        public TestModel(String name, String desc, int num) {
            this.name = name;
            this.Desc = desc;
            this.num = num;
        }

        public String getName() {
            return name;
        }

        public String getDesc() {
            return Desc;
        }

        public int getNum() {
            return num;
        }
    }
}
