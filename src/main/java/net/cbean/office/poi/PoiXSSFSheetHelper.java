package net.cbean.office.poi;

import net.cbean.office.poi.AbstractPoiSheetHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import static net.cbean.office.SheetHelper.DEFAULT_COLUMN_NAME_ROW;

public class PoiXSSFSheetHelper extends AbstractPoiSheetHelper {

    private XSSFSheet sheet;

    public PoiXSSFSheetHelper(XSSFSheet sheet) {
        this(sheet, DEFAULT_COLUMN_NAME_ROW);
    }

    public PoiXSSFSheetHelper(XSSFSheet sheet, CellParser cellParser) {
        this(sheet, DEFAULT_COLUMN_NAME_ROW, cellParser);
    }

    public PoiXSSFSheetHelper(XSSFSheet sheet, int columnNameRow) {
        this(sheet, columnNameRow, null);
    }

    public PoiXSSFSheetHelper(XSSFSheet sheet, int columnNameRow, CellParser cellParser) {
        this.sheet = sheet;
        super.init(sheet, columnNameRow, cellParser);
    }

    @Override
    Row getRow(int rowNum) {
        return this.sheet.getRow(rowNum);
    }
}
