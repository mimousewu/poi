package net.cbean.office.poi;

import net.cbean.office.SheetHelper;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.*;

import java.util.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public abstract class AbstractPoiSheetHelper implements SheetHelper {

    protected Map<String, Integer> columns = new HashMap<>();
    protected int rowCapacity;
    protected CellParser cellParser;

    /**
     * line number in Excel files. start from 0
     */
    protected int columnNameRow;

    protected void init(Sheet sheet, int columnNameRow, CellParser cellParser) {
        if (sheet == null) {
            throw new IllegalArgumentException("sheet not exists!");
        }
        this.cellParser = Optional.ofNullable(cellParser).orElse(new CellParser() {
            @Override
            public String parse(Cell cell) {
                String value = null;
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    value = new CellDateFormatter(cell.getCellStyle().getDataFormatString()).format(date);
                } else {
                    cell.setCellType(CellType.STRING);
                    value = cell.getStringCellValue();
                }
                return value;
            }
        });

        findHeadRow(sheet, columnNameRow);
    }

    private void findHeadRow(Sheet sheet, int columnNameRow) {
        columnNameRow = columnNameRow < sheet.getFirstRowNum() ? sheet.getFirstRowNum() : columnNameRow;
        Row headRow = sheet.getRow(columnNameRow);
        this.rowCapacity = sheet.getLastRowNum();

        for (int i = columnNameRow; i <= this.rowCapacity; i++) {
            if ((headRow = sheet.getRow(i)) != null) {
                this.columnNameRow = i;
                break;
            }
        }

        int columnCount = headRow.getLastCellNum();
        for (int i = 0; i < columnCount; i++) {
            Cell cell = headRow.getCell(i);
            if (cell != null)
                columns.put(this.cellParser.parse(cell), i);
        }
    }

    @Override
    public int getRows() {
        return this.rowCapacity - columnNameRow;
    }

    @Override
    public boolean containsColumn(String colName) {
        return columns.containsKey(colName);
    }

//    private ReadWriteLock locker = new ReentrantReadWriteLock();

    @Override
    public String getValue(String colName, int rowIndex) {
//        locker.readLock().lock();
//        try {
        Row row = getRow(getRealRowIndex(rowIndex));
        if (row == null) {
            return null;
        }

        if (columns.get(colName) == null) {
            throw new IllegalArgumentException("ColumnName not found! " + colName);
        }
        Cell cell = row.getCell(columns.get(colName).shortValue());
        if (cell == null) {
            return null;
        }
        return this.cellParser.parse(cell);
//        } finally {
//            locker.readLock().unlock();
//        }

    }

    private int getRealRowIndex(int rowIndex) {
        return rowIndex + columnNameRow + 1;
    }

    @Override
    public Iterator<Map<String, String>> iterator() {
        return new Iterator<Map<String, String>>() {
            private int position = 0;

            public boolean hasNext() {
                return position < getRows();
            }

            public Map<String, String> next() {
                Map<String, String> row = new HashMap<>();
                for (String columnName : columns.keySet()) {
                    row.put(columnName, getValue(columnName, position));
                }
                position++;
                return row;
            }

            public void remove() {
                throw new IllegalStateException(
                        "can't remove row from Excel sheet");
            }

        };
    }

    @Override
    public Map<String, String> getRowValue(int rowIndex) {
        Map<String, String> row = new HashMap<>();
        for (String columnName : columns.keySet()) {
            row.put(columnName, getValue(columnName, rowIndex));
        }
        return row;
    }

    @Override
    public String[] getColumns() {
        return this.columns.keySet().toArray(new String[]{});
    }

    @Override
    public String[] rowToArray(Map<String, String> rowValue, String... columnName) {
        columnName = columnName.length == 0 ? this.getColumns() : columnName;
        String[] result = new String[columnName.length];
        for (int i = 0; i < columnName.length; i++) {
            result[i] = rowValue.get(columnName[i]);
        }
        return result;
    }

    @Override
    public void iterate(Visitor visitor) {
        this.iterate(visitor, 0, this.rowCapacity - this.columnNameRow);
    }

    @Override
    public boolean iterate(Visitor visitor, int start, int offset) {
        if (this.rowCapacity - this.columnNameRow < start) {
            return false;
        }

        int end = start + offset > this.rowCapacity - this.columnNameRow ? this.rowCapacity - this.columnNameRow : start + offset;
        for (int i = start; i < end; i++) {
            visitor.visit(this.getRowValue(i), i);
        }

        return this.rowCapacity - this.columnNameRow > end;
    }

    abstract Row getRow(int rowNum);

    @Override
    public Stream<Map<String, String>> stream() {
        Iterator<Map<String, String>> iterator = iterator();

        return StreamSupport.stream(new Iterable<Map<String, String>>() {
            @Override
            public Iterator<Map<String, String>> iterator() {
                return iterator;
            }
        }.spliterator(), false);
    }
}
