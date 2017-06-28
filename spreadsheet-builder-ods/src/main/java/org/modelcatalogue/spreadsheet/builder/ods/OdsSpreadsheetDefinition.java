package org.modelcatalogue.spreadsheet.builder.ods;

import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.modelcatalogue.spreadsheet.builder.api.*;
import org.modelcatalogue.spreadsheet.impl.AbstractSheetDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractWorkbookDefinition;

import java.io.File;
import java.io.OutputStream;

class OdsSpreadsheetDefinition extends AbstractWorkbookDefinition implements SpreadsheetDefinition, WorkbookDefinition {

    private final SpreadSheet spreadsheet;

    OdsSpreadsheetDefinition(SpreadSheet spreadsheet) {
        this.spreadsheet = spreadsheet;
    }

    @Override
    public void writeTo(OutputStream outputStream) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    public void writeTo(File file) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected CellStyleDefinition createCellStyle() {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected AbstractSheetDefinition createSheet(String name) {
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
