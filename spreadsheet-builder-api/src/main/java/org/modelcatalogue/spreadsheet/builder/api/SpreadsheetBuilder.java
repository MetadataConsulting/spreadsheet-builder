package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.io.File;
import java.io.InputStream;

public interface SpreadsheetBuilder {

    SpreadsheetDefinition build(@DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition);
    SpreadsheetDefinition build(InputStream template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition);
    SpreadsheetDefinition build(File template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition);

}
