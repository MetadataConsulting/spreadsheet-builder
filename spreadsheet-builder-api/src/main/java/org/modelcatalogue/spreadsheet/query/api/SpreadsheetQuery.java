package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Cell;

import java.io.File;
import java.io.InputStream;
import java.util.Collection;


/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface SpreadsheetQuery {

    Collection<Cell> query(File spreadsheet);
    Collection<Cell> query(File spreadsheet, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookCriterion") Closure workbookCriterion);
    Collection<Cell> query(InputStream inputStream);
    Collection<Cell> query(InputStream inputStream, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookCriterion") Closure workbookCriterion);

}
