package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.io.File;
import java.io.OutputStream;

public interface SpreadsheetBuilder {

    void build(OutputStream outputStream, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition);
    void build(File file, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition);


}
