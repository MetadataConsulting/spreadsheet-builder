package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;

import java.io.FileNotFoundException;
import java.util.function.Consumer;

/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface SpreadsheetCriteria {

    SpreadsheetCriteriaResult all();
    SpreadsheetCriteriaResult query(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    Cell find(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    boolean exists(Consumer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;

}
