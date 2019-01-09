package net.cbean.office;

import java.util.Iterator;
import java.util.stream.Stream;

public interface RowData {

    String getValue(int rowIndex, int colIndex);

    Stream<String[]> stream();

    Iterator<String[]> iterator();

    Stream<CellData> cellStream();

    Iterator<CellData> cellIterator();

    interface CellData {
        String data(int start, int offset);
        String data(String colName);
    }
}
