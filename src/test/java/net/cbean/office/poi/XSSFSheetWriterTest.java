package net.cbean.office.poi;

import org.apache.poi.ss.usermodel.Sheet;
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

    @Before
    public void setup() {
        wb = new SXSSFWorkbook();

        String[] headers = new String[]{"Name", "desc", "Num"};
        SXSSFSheet sh = wb.createSheet("sheet1");
        sheetWriter = new XSSFSheetWriter(sh, headers).init();
    }

    @After
    public void tearDown() throws Exception {
        String outputExcel = "target/test.xlsx";
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
        assertOutputExcel(outputExcel);
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
    public void testAddDataWithValueHandler() throws Exception {
        for (Object[] row : initData()) {
            sheetWriter.addRow((header, index) -> row[index]);
        }
    }

    @Test
    public void testAddDataListObject() throws Exception {
        List<TestModel> data = new ArrayList<>();
        data.add(new TestModel(new String("a"), new String("b"), 1));
        data.add(new TestModel(new String("a1"), new String("b1"), 2));
        data.add(new TestModel(new String("a2"), new String("b2"), 3));

        sheetWriter.addData(data);
    }

    @Test
    public void testAddDataList() throws Exception {
        sheetWriter.addData(initData());
    }

    @Test
    public void testAddDataArray() throws Exception {
        sheetWriter.addData(initData().toArray(new Object[]{}));
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
