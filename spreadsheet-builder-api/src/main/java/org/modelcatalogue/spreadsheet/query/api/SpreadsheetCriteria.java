package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

import java.io.FileNotFoundException;


/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface SpreadsheetCriteria {

    SpreadsheetCriteriaResult all();
    SpreadsheetCriteriaResult query(Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    Cell find(Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    boolean exists(Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;

}
