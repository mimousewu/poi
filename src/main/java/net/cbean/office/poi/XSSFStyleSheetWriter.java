package net.cbean.office.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class XSSFStyleSheetWriter extends XSSFSheetWriter {
    private List<XSSFCellStyle> headerStyles = new ArrayList<>();

    public XSSFStyleSheetWriter(SXSSFSheet sh, String[] header) {
        super(sh, header);
    }

    public XSSFStyleSheetWriter(SXSSFSheet sh, String[] header, String[] headerLabel, File styleTemplate) {
        this.sheet = sh;
        this.headers = header;
        if(header.length != headerLabel.length) {
            throw new IllegalArgumentException("Size of headerLabel is not match the size of header!");
        }

        if (styleTemplate != null) {
            loadStyle(styleTemplate);
        }

        //Override super init but need more parameters
        this.sheet.setRandomAccessWindowSize(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
        Row row = this.sheet.createRow(nextRowIndex());
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headerLabel[i]);
            if (headerStyles.size() > 0) {
                cell.setCellStyle(this.headerStyles.get(i));
            }
        }
    }

    private void loadStyle(File styleTemplate) {
        try {
            FileInputStream in = new FileInputStream(styleTemplate);
            Sheet tempSheet = new XSSFWorkbook(in).getSheetAt(0);
            for (int i = 0; i < this.headers.length; i++) {
                headerStyles.add((XSSFCellStyle) tempSheet.getRow(0).getCell(i).getCellStyle());
                CellStyle cellStyle = tempSheet.getRow(1).getCell(i).getCellStyle();
                super.setCellStyleWriter(this.headers[i], cell -> cell.setCellStyle(cellStyle));
            }
            in.close();
        } catch (IOException e) {
            throw new IllegalArgumentException("Failed to load style form file: "
                    + styleTemplate.getAbsolutePath(), e);
        }
    }
}
