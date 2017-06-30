package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition;
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractSheetDefinition;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

class PoiSheetDefinition extends AbstractSheetDefinition implements SheetDefinition, Sheet {

    private final XSSFSheet xssfSheet;

    PoiSheetDefinition(PoiWorkbookDefinition workbook, XSSFSheet xssfSheet) {
        super(workbook);
        this.xssfSheet = xssfSheet;
    }

    @Override
    protected PoiRowDefinition createRow(int zeroBasedRowNumber) {
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

    public List<PoiRowDefinition> getRows() {
        List<PoiRowDefinition> rows = new ArrayList<PoiRowDefinition>();
        for (Row it : xssfSheet) {
            PoiRowDefinition row = getRowByNumber(it.getRowNum() + 1);
            if (row != null) {
                rows.add(row);
            }
            rows.add(createRowWrapper(it.getRowNum() + 1));

        }
        return Collections.unmodifiableList(rows);
    }

    PoiRowDefinition getRowByNumber(int rowNumberStartingOne) {
        return (PoiRowDefinition) this.rows.get(rowNumberStartingOne);
    }

    @Override
    public Sheet getNext() {
        int current = getWorkbook().getWorkbook().getSheetIndex(xssfSheet.getSheetName());

        if (current == getWorkbook().getWorkbook().getNumberOfSheets() - 1) {
            return null;
        }
        XSSFSheet next = getWorkbook().getWorkbook().getSheetAt(current + 1);

        for (Sheet sheet : getWorkbook().getSheets()) {
            if (sheet.getName().equals(next.getSheetName())) {
                return sheet;
            }
        }


        return new PoiSheetDefinition(getWorkbook(), next);
    }

    @Override
    public Sheet getPrevious() {
        int current = getWorkbook().getWorkbook().getSheetIndex(xssfSheet.getSheetName());

        if (current == 0) {
            return null;
        }
        XSSFSheet next = getWorkbook().getWorkbook().getSheetAt(current - 1);
        for (Sheet sheet : getWorkbook().getSheets()) {
            if (sheet.getName().equals(next.getSheetName())) {
                return sheet;
            }
        }
        return new PoiSheetDefinition(getWorkbook(), next);
    }

    @Override
    public Page getPage() {
        return new PoiPageSettingsProvider(this);
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

    private PoiRowDefinition createRowWrapper(int oneBasedRowNumber) {
        PoiRowDefinition newRow = new PoiRowDefinition(this, xssfSheet.getRow(oneBasedRowNumber - 1));
        rows.put(oneBasedRowNumber, newRow);
        return newRow;
    }

    public <T> T asType(Class<T> type) {
        if (type.isInstance(xssfSheet)) {
            return type.cast(xssfSheet);
        }
        return DefaultGroovyMethods.asType(this, type);
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
