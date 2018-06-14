package net.cbean.office.poi;

import org.apache.poi.hssf.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiHSSFSheetHelperTest extends AbstractPoiSheetHelperTest {

    @BeforeClass
    public static void startup() throws Exception {
        FileOutputStream outFile = new FileOutputStream("target/out.xls");

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("exportsheet");
        try {
            addCell(sheet, 3, 0, "Column1");
            addCell(sheet, 4, 0, "Value 1");
            addCell(sheet, 5, 0, "Value11");
            addCell(sheet, 6, 0, "Value111");
            addCell(sheet, 3, 1, "Column2");
            addCell(sheet, 4, 1, "Value2");
            addCell(sheet, 5, 1, "Value22");
            addCell(sheet, 6, 1, "Value222");
            addCell(sheet, 7, 1, "Value2222");

            // There is issue that excel can only store the last column
            workbook.write(outFile);
            outFile.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outFile.close();
        }
    }

    @Before
    public void setup() throws Exception {
        in = new FileInputStream("target/out.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(in);
        helper = new PoiHSSFSheetHelper(workbook.getSheetAt(0));
    }

    private static void addCell(HSSFSheet sheet, int rowc, int colc, String content) {
        HSSFRow row = sheet.createRow(rowc);
        HSSFCell cell = row.createCell(colc);
        HSSFRichTextString str = new HSSFRichTextString(content);
        cell.setCellValue(str);
    }

    @After
    public void tearDown() throws Exception {
        in.close();
    }
}
