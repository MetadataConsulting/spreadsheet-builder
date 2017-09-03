package builders.dsl.spreadsheet.api;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * Wraps row so it can be accessible using the names from headers.
 */
public final class DataRow implements Row {

    public static DataRow create(Row row, Row headersRow) {
        Collection<? extends Cell> cells = headersRow.getCells();
        Map<String, Integer> mapping = new HashMap<String, Integer>(cells.size());
        for (Cell cell : cells) {
            mapping.put(String.valueOf(cell.getValue()), cell.getColumn());
        }
        return new DataRow(row, mapping);
    }

    public static DataRow create(Map<String, Integer> mapping, Row row) {
        return new DataRow(row, new HashMap<String, Integer>(mapping));
    }

    private final Row row;
    private final Map<String, Cell> cells;

    private DataRow(Row row, Map<String, Integer> mapping) {
        this.row = row;

        Map<String, Cell> cells = new HashMap<String, Cell>(row.getCells().size());

        for (Map.Entry<String, Integer> entry : mapping.entrySet()) {
            for (Cell cell : row.getCells()) {
                if (cell.getColumn() == entry.getValue()) {
                    cells.put(entry.getKey(), cell);
                }
            }
        }

        this.cells = cells;
    }

    /**
     * Returns the cell by the label given from headers or mapping.
     * @param name label from header or mapping
     * @return the cell by the label given from headers or mapping
     */
    public Cell get(String name) {
        return cells.get(name);
    }

    @Override
    public int getNumber() {
        return row.getNumber();
    }

    @Override
    public Sheet getSheet() {
        return row.getSheet();
    }

    @Override
    public Collection<? extends Cell> getCells() {
        return row.getCells();
    }

    @Override
    public Row getAbove() {
        return row.getAbove();
    }

    @Override
    public Row getAbove(int howMany) {
        return row.getAbove(howMany);
    }

    @Override
    public Row getBellow() {
        return row.getBellow();
    }

    @Override
    public Row getBellow(int howMany) {
        return row.getBellow(howMany);
    }
}
