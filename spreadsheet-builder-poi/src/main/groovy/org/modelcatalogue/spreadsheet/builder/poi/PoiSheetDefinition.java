package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition;
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractSheetDefinition;

class PoiSheetDefinition extends AbstractSheetDefinition implements SheetDefinition {

    private final XSSFSheet xssfSheet;

    PoiSheetDefinition(PoiWorkbookDefinition workbook, XSSFSheet xssfSheet) {
        super(workbook);
        this.xssfSheet = xssfSheet;
    }

    // TODO: make protected again
    @Override public PoiRowDefinition createRow(int zeroBasedRowNumber) {
        XSSFRow row = xssfSheet.getRow(zeroBasedRowNumber);

        if (row == null) {
            row = xssfSheet.createRow(zeroBasedRowNumber);
        }

        return new PoiRowDefinition(this, row);
    }

    @Override
    protected PageDefinition createPageDefintion() {
        return new PoiPageSettingsProvider(this);
    }

    @Override
    protected void applyRowGroup(int startPosition, int endPosition, boolean collapsed) {
        xssfSheet.groupRow(startPosition, endPosition);
        if (collapsed) {
            xssfSheet.setRowGroupCollapsed(endPosition, true);
        }
    }

    @Override
    public String getName() {
        return xssfSheet.getSheetName();
    }

    @Override
    public PoiWorkbookDefinition getWorkbook() {
        return (PoiWorkbookDefinition) super.getWorkbook();
    }

    // TODO: make protected again
    public PoiRowDefinition getRowByNumber(int rowNumberStartingOne) {
        return (PoiRowDefinition) this.rows.get(rowNumberStartingOne);
    }

    @Override
    protected void doFreeze(int column, int row) {
        xssfSheet.createFreezePane(column, row);
    }

    @Override
    protected void doLock() {
        xssfSheet.enableLocking();
    }

    @Override
    protected void doPassword(String password) {
        xssfSheet.protectSheet(password);
    }

    protected XSSFSheet getSheet() {
        return xssfSheet;
    }

    protected void processAutoColumns() {
        for (Integer index : autoColumns) {
            xssfSheet.autoSizeColumn(index);
        }
    }

    protected void processAutomaticFilter() {
        if (automaticFilter && xssfSheet.getLastRowNum() > 0) {
            xssfSheet.setAutoFilter(new CellRangeAddress(
                    xssfSheet.getFirstRowNum(),
                    xssfSheet.getLastRowNum(),
                    xssfSheet.getRow(xssfSheet.getFirstRowNum()).getFirstCellNum(),
                    xssfSheet.getRow(xssfSheet.getFirstRowNum()).getLastCellNum() - 1
            ));
        }
    }
}
