package net.cbean.office;

import org.apache.poi.ss.usermodel.Cell;

import java.util.List;

public interface SheetWriter {

    /**
     * Client decide how to deal with each lines of row.
     *
     * @param rowHandler handler value of rows
     */
    void addRow(RowHandler rowHandler);

    /**
     * This will automatic check the entry of list and use row handler to deal with values
     *
     * @param data if the entry of list in an object, it will use each bean property as cell values,
     *             the toString() method will be called if it is not a string property
     * @param <T> Any type of row value will be supported
     */
    <T> void addData(List<T> data);

    /**
     * This is the same as above method, but use array instead of List.
     * @param data most of time this parameter will be double-dimension array like String[][]
     * @param <T> cloud be array or object
     */
    <T> void addData(T[] data);

    @FunctionalInterface
    interface RowHandler {
        Object cellValue(String header, int colIndex);
    }

    @FunctionalInterface
    interface CellStyleWriter {
        void handleCellStyle(Cell cell);
    }
}
