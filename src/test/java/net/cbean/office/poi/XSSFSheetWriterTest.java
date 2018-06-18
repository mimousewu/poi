package net.cbean.office.poi;

import net.cbean.office.poi.XSSFSheetWriter;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
        sheetWriter = new XSSFSheetWriter(sh, headers);
    }

    @After
    public void tearDown() {
        try {
            fos = new FileOutputStream("sxssf.xlsx");
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
        assertTrue(false);
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
