package net.cbean.office.poi;

import org.apache.poi.xssf.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiXSSFSheetHelperTest extends AbstractPoiSheetHelperTest {

    @BeforeClass
    public static void startup() throws Exception {
        FileOutputStream outFile = new FileOutputStream("target/out.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("exportsheet");
        try {
            addCell(sheet, 0, 0, "Column1");
            addCell(sheet, 1, 0, "Value 1");
            addCell(sheet, 2, 0, "Value11");
            addCell(sheet, 3, 0, "Value111");
            addCell(sheet, 0, 1, "Column2");
            addCell(sheet, 1, 1, "Value2");
            addCell(sheet, 2, 1, "Value22");
            addCell(sheet, 3, 1, "Value222");
            addCell(sheet, 4, 1, "Value2222");
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
        in = new FileInputStream("target/out.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        helper = new PoiXSSFSheetHelper(workbook.getSheetAt(0));
    }

    private static void addCell(XSSFSheet sheet, int rowc, int colc, String content) {
        XSSFRow row = sheet.createRow(rowc);
        XSSFCell cell = row.createCell(colc);
        XSSFRichTextString str = new XSSFRichTextString(content);
        cell.setCellValue(str);
    }

    @After
    public void tearDown() throws Exception {
        in.close();
    }
}
