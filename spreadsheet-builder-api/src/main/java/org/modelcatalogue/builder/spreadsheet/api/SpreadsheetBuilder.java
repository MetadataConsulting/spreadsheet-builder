package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.io.OutputStream;

public interface SpreadsheetBuilder {

    void build(OutputStream outputStream, @DelegatesTo(Workbook.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Workbook") Closure workbookDefinition);


}
