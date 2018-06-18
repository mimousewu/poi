package net.cbean.office.poi;

import net.cbean.office.SheetWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Array;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

public class XSSFSheetWriter implements SheetWriter {
    private SXSSFSheet sheet;
    private int rowIndex = 0;
    private String[] headers;

    /**
     * key: header
     */
    private Map<String, CellStyleWriter> cellStyleWriters = new HashMap<>();

    public XSSFSheetWriter(SXSSFSheet sh, String[] header) {
        this.sheet = sh;
        this.headers = header;
        init();
    }

    private void init() {
        this.sheet.setRandomAccessWindowSize(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
        Row row = this.sheet.createRow(nextRowIndex());
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
        }
    }

    public void setCellStyleWriter(String header, CellStyleWriter cellStyleWriter) {
        this.cellStyleWriters.put(header, cellStyleWriter);
    }

    public void addRow(RowHandler valueHandler) {
        Row row = this.sheet.createRow(nextRowIndex());
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            String header = headers[i];
            if (this.cellStyleWriters.containsKey(header)) {
                this.cellStyleWriters.get(header).handleCellStyle(cell);
            }
            fillValue(cell, valueHandler.cellValue(header, i));
        }
    }

    @Override
    public <T> void addData(List<T> data) {
        data.forEach(rowData -> {
            addRowData(rowData);
        });
    }

    @Override
    public <T> void addData(T[] data) {
        for (int i = 0; i < data.length; i++) {
            addRowData(data[i]);
        }
    }

    private <T> void addRowData(T rowData) {
        if (rowData != null && rowData.getClass().isArray()) {
            this.addRow((header, index) -> Array.get(rowData, index));
        } else {
            Class<T> clazz = (Class<T>) rowData.getClass();
            this.addRow((header, index) -> {
                return getObjectCellValue(rowData, clazz, header);
            });
        }
    }

    private <T> Object getObjectCellValue(T rowData, Class<T> clazz, String header) {
        try {
            Method method = null;
            try {
                method = clazz.getMethod("get" + header);
            } catch (NoSuchMethodException nme) {
                method = clazz.getMethod("get" + capitalizeInitialLetter(header));
            }
            return method.invoke(rowData, (Object[]) null);
        } catch (IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new IllegalArgumentException("Failed in get [" + header + "] Property from class "
                    + clazz.getName(), e);
        }
    }

    private void fillValue(Cell cell, Object value) {
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private int nextRowIndex() {
        return this.rowIndex++;
    }

    private static final String capitalizeInitialLetter(String name) {
        String firstCapital = name.substring(0, 1);
        return firstCapital.toUpperCase() + name.substring(1);
    }
}
