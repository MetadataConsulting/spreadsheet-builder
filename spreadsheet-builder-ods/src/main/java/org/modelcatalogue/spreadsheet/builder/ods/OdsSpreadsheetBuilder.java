package org.modelcatalogue.spreadsheet.builder.ods;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder;
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition;
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition;

import java.io.File;
import java.io.InputStream;

public enum OdsSpreadsheetBuilder implements SpreadsheetBuilder {

    INSTANCE;


    @Override
    public SpreadsheetDefinition build(Configurer<WorkbookDefinition> workbookDefinition) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    public SpreadsheetDefinition build(InputStream template, Configurer<WorkbookDefinition> workbookDefinition) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    public SpreadsheetDefinition build(File template, Configurer<WorkbookDefinition> workbookDefinition) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

}
