package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Configurer;

import java.io.File;
import java.io.InputStream;

public interface SpreadsheetBuilder {

    SpreadsheetDefinition build(Configurer<WorkbookDefinition> workbookDefinition);
    SpreadsheetDefinition build(InputStream template, Configurer<WorkbookDefinition> workbookDefinition);
    SpreadsheetDefinition build(File template, Configurer<WorkbookDefinition> workbookDefinition);

}
