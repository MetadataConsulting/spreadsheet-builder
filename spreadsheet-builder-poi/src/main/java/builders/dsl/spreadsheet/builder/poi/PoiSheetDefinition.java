package builders.dsl.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import builders.dsl.spreadsheet.builder.api.PageDefinition;
import builders.dsl.spreadsheet.builder.api.SheetDefinition;
import builders.dsl.spreadsheet.impl.AbstractSheetDefinition;

class PoiSheetDefinition extends AbstractSheetDefinition implements SheetDefinition {

    private static final int WIDTH_ARROW_BUTTON = 2 * 255;
    public static final int MAX_COLUMN_WIDTH = 255 * 256;

    private final XSSFSheet xssfSheet;

    PoiSheetDefinition(PoiWorkbookDefinition workbook, XSSFSheet xssfSheet) {
        super(workbook);
        this.xssfSheet = xssfSheet;
    }

    @Override protected PoiRowDefinition createRow(int zeroBasedRowNumber) {
        XSSFRow row = xssfSheet.getRow(zeroBasedRowNumber);

        if (row == null) {
            row = xssfSheet.createRow(zeroBasedRowNumber);
        }

        return new PoiRowDefinition(this, row);
    }

    @Override
    protected PageDefinition createPageDefinition() {
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

    @Override
    protected void doFreeze(int column, int row) {
        xssfSheet.createFreezePane(column, row);
    }

    @Override
    protected void doLock() {
        xssfSheet.enableLocking();
    }

    @Override
    protected void doHide() {
        setVisibility(SheetVisibility.HIDDEN);
    }

    private void setVisibility(SheetVisibility visibility) {
        xssfSheet.getWorkbook().setSheetVisibility(xssfSheet.getWorkbook().getSheetIndex(xssfSheet), visibility);
    }

    @Override
    protected void doHideCompletely() {
        setVisibility(SheetVisibility.VERY_HIDDEN);
    }

    @Override
    protected void doShow() {
        setVisibility(SheetVisibility.VISIBLE);
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
            if (automaticFilter) {
                xssfSheet.setColumnWidth(index, Math.min(xssfSheet.getColumnWidth(index) + WIDTH_ARROW_BUTTON, MAX_COLUMN_WIDTH));
            }
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
