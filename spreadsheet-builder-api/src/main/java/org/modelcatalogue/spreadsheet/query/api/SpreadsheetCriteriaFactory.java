package org.modelcatalogue.spreadsheet.query.api;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.InputStream;


/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface SpreadsheetCriteriaFactory {

    SpreadsheetCriteria forFile(File spreadsheet) throws FileNotFoundException;
    SpreadsheetCriteria forStream(InputStream inputStream);

}
