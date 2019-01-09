package net.cbean.office.poi;

import net.cbean.office.SheetHelper;
import org.junit.Ignore;
import org.junit.Test;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.*;
import java.util.concurrent.locks.ReadWriteLock;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import static org.junit.Assert.*;

@Ignore
public class AbstractPoiSheetHelperTest {
    protected InputStream in;
    protected SheetHelper helper;


    @Test
    public void testGetRows() throws Exception {
        assertEquals(4, helper.getRows());
    }

    @Test
    public void testContainsColumn() throws Exception {
        assertTrue(helper.containsColumn("Column2"));
    }

    @Test
    public void testIterator() throws Exception {
        for (Iterator<Map<String, String>> iter = helper.iterator(); iter
                .hasNext(); ) {
            Map<String, String> entry = iter.next();
            assertNotNull(entry.get("Column2"));
        }
    }

    @Test
    public void testGetValue() throws Exception {
        assertEquals("Value2", helper.getValue("Column2", 0));
        assertEquals("Value22", helper.getValue("Column2", 1));
    }

    @Test
    public void testRowDataGetValue() throws Exception {
        assertEquals("Value2", helper.rowData().getValue(0,0));
        assertEquals("Value22", helper.rowData().getValue(1,0));
    }

    @Test
    public void testGetRowValue() throws Exception {
        assertEquals("Value2", helper.getRowValue(0).get("Column2"));
        assertEquals("Value22", helper.getRowValue(1).get("Column2"));
    }

    @Test
    public void testGetColumns() throws Exception {
        String[] columns = helper.getColumns();
        assertEquals("Column2", columns[0]);
    }

    @Test
    public void testRowToArrayy() throws Exception {
        assertEquals("Value2", helper.rowToArray(helper.getRowValue(0))[0]);
        assertEquals("Value22", helper.rowToArray(helper.getRowValue(1))[0]);
    }

    @Test
    public void testIterate() throws Exception {
        List<String[]> out = new ArrayList<>();
        helper.iterate((row, index) -> {
            out.add(helper.rowToArray(row));
        });

        assertEquals(4, out.size());
        assertEquals("Value2", out.get(0)[0]);
    }

    @Test
    public void testIterateWithStartOffset() throws Exception {
        List<String[]> out = new ArrayList<>();
        int offset = 2;
        for (int s = 1; helper.iterate((row, index) -> {
            out.add(helper.rowToArray(row));
        }, s, offset); s += offset)
            ;

        assertEquals(3, out.size());
        assertEquals("Value22", out.get(0)[0]);
    }

    @Test
    public void testIterateinMultipleThread() throws Exception {
        ExecutorService executor = Executors.newFixedThreadPool(10);
        List<Future<Boolean>> futures = new ArrayList<>();

        List<String[]> out = new ArrayList<>();
        ReadWriteLock locker = new ReentrantReadWriteLock();

        int offset = 2;
        for (int i = 1; i < helper.getRows(); i += offset) {

            Callable callable = new Callable<Boolean>() {
                private int start;

                public Callable init(int s) {
                    this.start = s;
                    return this;
                }

                @Override
                public Boolean call() throws Exception {

                    return helper.iterate((row, index) -> {
                        locker.writeLock().lock();
                        try {
                            out.add(helper.rowToArray(row));
                        } finally {
                            locker.writeLock().unlock();
                        }

                    }, this.start, offset);
                }
            }.init(i);

            futures.add(executor.submit(callable));

        }

        futures.forEach(future -> {
            try {
                future.get();
            } catch (InterruptedException | ExecutionException e) {
                e.printStackTrace();
            }
        });

        executor.shutdown();

        assertEquals(3, out.size());
        assertNotEquals(null, out.get(0)[0]);
        assertNotEquals(null, out.get(1)[0]);
        assertNotEquals(null, out.get(2)[0]);
    }

    @Test
    public void testStream() throws Exception {
        assertFalse(helper.stream().anyMatch(row -> row.isEmpty()));
    }

    @Test
    public void testCellStream() throws Exception {
        String value = helper.rowData().cellStream()
                .findFirst().map(cellData -> cellData.data(1, 0)).orElse(null);
        assertEquals("Value2", value);

        value = helper.rowData().cellStream()
                .findFirst().map(cellData -> cellData.data("Column2")).orElse(null);
        assertEquals("Value22", value);
    }
}
