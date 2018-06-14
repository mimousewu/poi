package net.cbean.office;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Iterator;
import java.util.Map;
import java.util.stream.Stream;

public interface SheetHelper {
    public static final int DEFAULT_COLUMN_NAME_ROW = 0;

    int getRows();

    boolean containsColumn(String colName);

    /**
     * @param colName
     * @param rowIndex start from 0
     * @return
     */
    String getValue(String colName, int rowIndex);

    Iterator<Map<String, String>> iterator();

    /**
     * @param rowIndex start from 0
     * @return
     */
    Map<String, String> getRowValue(int rowIndex);

    /**
     * get the columns during initialization
     *
     * @return
     */
    String[] getColumns();

    /**
     * get row values by given columns, if column not given, use the columns during initialization
     *
     * @param row
     * @param columnName
     * @return
     */
    String[] rowToArray(Map<String, String> rowValue, String... columnName);

    Stream<Map<String, String>> stream();

    void iterate(Visitor visitor);

    /**
     * @param visitor
     * @param start
     * @param offset
     * @return if there are more rows, return true
     */
    boolean iterate(Visitor visitor, int start, int offset);

    @FunctionalInterface
    public interface Visitor {
        void visit(Map<String, String> row, Integer rowIndex);
    }

    @FunctionalInterface
    public interface CellParser {
        String parse(Cell cell);
    }
}
