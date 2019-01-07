package net.cbean.office.poi;

import net.cbean.office.RowData;
import org.apache.poi.ss.usermodel.Row;

import java.util.Iterator;
import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class RowDataHelper implements RowData {
    AbstractPoiSheetHelper sheetHelper;

    public RowDataHelper(AbstractPoiSheetHelper sheetHelper) {
        this.sheetHelper = sheetHelper;
    }

    @Override
    public String getValue(int rowIndex, int colIndex) {
        Row row = sheetHelper.getRow(sheetHelper.getRealRowIndex(rowIndex));
        if (row == null) {
            return null;
        }

        return Optional.ofNullable(row.getCell(colIndex)).map(
                c -> sheetHelper.cellParser.parse(c)).orElse(null);
    }

    @Override
    public Stream<String[]> stream() {
        Iterator<String[]> iterator = iterator();

        return StreamSupport.stream((
                (Iterable<String[]>) () -> iterator)
                .spliterator(), false);
    }

    @Override
    public Iterator<String[]> iterator() {
        return new Iterator<String[]>() {
            private int position = 0;

            public boolean hasNext() {
                return position < sheetHelper.getRows();
            }

            public String[] next() {
                String[] data = new String[sheetHelper.rowCapacity];
                for (int i = 0; i < data.length; i++) {
                    data[i] = getValue(position, i);
                }

                position++;
                return data;
            }

            public void remove() {
                throw new IllegalStateException(
                        "can't remove row from Excel sheet");
            }

        };
    }
}
