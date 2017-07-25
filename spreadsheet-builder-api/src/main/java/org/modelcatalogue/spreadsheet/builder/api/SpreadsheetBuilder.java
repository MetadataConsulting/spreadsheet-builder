package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Configurer;

import java.io.File;
import java.io.InputStream;

public interface SpreadsheetBuilder {
    void build(Configurer<WorkbookDefinition> workbookDefinition);
}
