package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;

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
