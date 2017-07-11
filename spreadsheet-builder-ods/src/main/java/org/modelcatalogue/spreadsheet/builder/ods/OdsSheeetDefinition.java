package org.modelcatalogue.spreadsheet.builder.ods;

import org.jopendocument.dom.spreadsheet.Sheet;
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition;
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractRowDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractSheetDefinition;

class OdsSheeetDefinition extends AbstractSheetDefinition {

    private final Sheet sheet;

    OdsSheeetDefinition(OdsSpreadsheetDefinition workbook, Sheet sheet) {
        super(workbook);
        this.sheet = sheet;
    }

    @Override
    protected void processAutoColumns() {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected void processAutomaticFilter() {
        throw new UnsupportedOperationException("Not yet implemented");
    }


    @Override
    public AbstractRowDefinition createRow(int zeroBasedRowNumber) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected void doFreeze(int column, int row) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected PageDefinition createPageDefinition() {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected void applyRowGroup(int startPosition, int endPosition, boolean collapsed) {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected String getName() {
        return sheet.getName();
    }

    @Override
    protected void doLock() {
        throw new UnsupportedOperationException("Not yet implemented");
    }

    @Override
    protected void doPassword(String password) {
        throw new UnsupportedOperationException("Not yet implemented");
    }
}
