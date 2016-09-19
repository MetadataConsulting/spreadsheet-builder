package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Workbook;

import java.io.InputStream;

public interface WorkbookSupplier {

    Workbook load(InputStream stream);

}
