package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;

import java.util.Iterator;

public abstract class AbstractSpreadsheetCriteriaResult implements SpreadsheetCriteriaResult {

    @Override
    public Iterator<Cell> iterator() {
        return getCells().iterator();
    }

    @Override
    public String toString() {
        return getCells().toString();
    }
}
