package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.io.File;
import java.io.InputStream;
import java.util.Collection;


/**
 * Cell matcher uses the builder like syntax to find cells within the workbook.
 * Not all the constructs are be supported at the moment.
 * Check the documentation for the list of all supported features.
 */
public interface CellMatcher {

    Collection<CellDefinition> match(File spreadsheet, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookDefinition") Closure workbookDefinition);
    Collection<CellDefinition> match(InputStream inputStream, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookDefinition") Closure workbookDefinition);

}
