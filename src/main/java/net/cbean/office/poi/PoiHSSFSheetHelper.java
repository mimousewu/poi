package net.cbean.office.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;

import java.util.stream.Stream;

public class PoiHSSFSheetHelper extends AbstractPoiSheetHelper {

    private HSSFSheet sheet;

    public PoiHSSFSheetHelper(HSSFSheet sheet) {
        this(sheet, DEFAULT_COLUMN_NAME_ROW, null);
    }

    public PoiHSSFSheetHelper(HSSFSheet sheet, CellParser cellParser) {
        this(sheet, DEFAULT_COLUMN_NAME_ROW, cellParser);
    }

    public PoiHSSFSheetHelper(HSSFSheet sheet, int columnNameRow) {
        this(sheet, columnNameRow, null);
    }

    public PoiHSSFSheetHelper(HSSFSheet sheet, int columnNameRow, CellParser cellParser) {
        this.sheet = sheet;
        super.init(sheet, columnNameRow, cellParser);
    }

    @Override
    Row getRow(int rowNum) {
        return this.sheet.getRow(rowNum);
    }
}
