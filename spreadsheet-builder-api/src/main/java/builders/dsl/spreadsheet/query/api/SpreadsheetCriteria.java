package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Cell;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.io.FileNotFoundException;


/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface SpreadsheetCriteria {

    SpreadsheetCriteriaResult all();
    SpreadsheetCriteriaResult query(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.WorkbookCriterion") Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    Cell find(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.WorkbookCriterion") Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;
    boolean exists(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.WorkbookCriterion") Configurer<WorkbookCriterion> workbookCriterion) throws FileNotFoundException;

}
