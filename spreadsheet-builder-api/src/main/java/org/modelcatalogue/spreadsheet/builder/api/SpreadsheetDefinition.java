package org.modelcatalogue.spreadsheet.builder.api;

import java.io.File;
import java.io.OutputStream;

public interface SpreadsheetDefinition {

    void writeTo(OutputStream outputStream);
    void writeTo(File file);

}
