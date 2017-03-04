package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Sheet;

import java.util.Collection;

public interface SpreadsheetCriteriaResult extends Iterable<Cell> {

    Collection<Cell> getCells();
    Collection<Row> getRows();
    Collection<Sheet> getSheets();

}
