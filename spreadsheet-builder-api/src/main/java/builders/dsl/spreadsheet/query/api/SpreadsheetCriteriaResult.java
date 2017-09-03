package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.api.Sheet;

import java.util.Collection;

public interface SpreadsheetCriteriaResult extends Iterable<Cell> {

    /**
     * Returns all cells matching the criteria.
     * @return all cells matching the criteria
     */
    Collection<Cell> getCells();

    /**
     * Returns all rows matching the criteria. If any cell criteria is present, at least one cell in the row
     * must pass the test.
     * @return all rows matching the criteria
     */
    Collection<Row> getRows();

    /**
     * Returns all sheets matching the criteria. If any row or cell criteria is present at least one row (or cell)
     * must pass the test.
     * @return all the sheets matching the criteria.
     */
    Collection<Sheet> getSheets();

    /**
     * Returns first cell matching the criteria or null.
     * @return first cell matching the criteria or null
     */
    Cell getCell();

    /**
     * Returns first row matching the criteria or null.
     * @return first row matching the criteria or null
     */
    Row getRow();

    /**
     * Returns first sheet matching the criteria or null.
     * @return first sheet matching the criteria or null
     */
    Sheet getSheet();

}
