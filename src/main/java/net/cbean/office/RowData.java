package net.cbean.office;

import java.util.Iterator;
import java.util.stream.Stream;

public interface RowData {

    String getValue(int rowIndex, int colIndex);

    Stream<String[]> stream();

    Iterator<String[]> iterator();

    @FunctionalInterface
    interface CellData {
        String data(int offset);
    }
}
