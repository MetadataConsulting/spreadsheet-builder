package org.modelcatalogue.spreadsheet.builder.poi;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition;
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractRowDefinition;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

class PoiRowDefinition extends AbstractRowDefinition implements RowDefinition, Row {

    private final XSSFRow xssfRow;

    PoiRowDefinition(PoiSheetDefinition sheet, XSSFRow xssfRow) {
        super(sheet);
        this.xssfRow = xssfRow;
    }

    public List<PoiCellDefinition> getCells() {
        // TODO: reuse existing cells
        List<PoiCellDefinition> cells = new ArrayList<PoiCellDefinition>();
        for (Cell cell : xssfRow) {
            cells.add(new PoiCellDefinition(this, (XSSFCell) cell));
        }
        return Collections.unmodifiableList(cells);
    }

    @Override
    protected CellDefinition createCell(int zeroBasedCellNumber) {
        XSSFCell cell = xssfRow.getCell(zeroBasedCellNumber);

        if (cell == null) {
            cell = xssfRow.createCell(zeroBasedCellNumber);
        }

        return new PoiCellDefinition(this, cell);
    }

    @Override
    protected void handleSpans(CellDefinition cell) {
        if (cell instanceof PoiCellDefinition) {
            PoiCellDefinition poiCell = (PoiCellDefinition) cell;
            if (poiCell.getColspan() > 1 || poiCell.getRowspan() > 1) {
                xssfRow.getSheet().addMergedRegion(poiCell.getCellRangeAddress());
            }
        } else {
            throw new IllegalArgumentException("Unsupported cell: " + cell);
        }
    }

    @Override
    public int getNumber() {
        return xssfRow.getRowNum() + 1;
    }

    @Override
    public PoiSheetDefinition getSheet() {
        return (PoiSheetDefinition) sheet;
    }

    protected XSSFRow getRow() {
        return xssfRow;
    }


    @Override
    public PoiRowDefinition getAbove(int howMany) {
        return aboveOrBellow(-howMany);
    }

    @Override
    public PoiRowDefinition getAbove() {
        return getAbove(1);
    }

    @Override
    public PoiRowDefinition getBellow(int howMany) {
        return aboveOrBellow(howMany);
    }

    @Override
    public PoiRowDefinition getBellow() {
        return getBellow(1);
    }

    private PoiRowDefinition aboveOrBellow(int howMany) {
        if (xssfRow.getRowNum() + howMany < 0 || xssfRow.getRowNum() + howMany >  xssfRow.getSheet().getLastRowNum()) {
            return null;
        }
        PoiRowDefinition existing = (PoiRowDefinition) sheet.getRowByNumber(getNumber() + howMany);
        if (existing != null) {
            return existing;
        }
        return (PoiRowDefinition) sheet.createRow(getNumber() + howMany);
    }

    protected void doCreateGroup(int startPosition, int endPosition, boolean collapsed) {
        getSheet().getSheet().groupColumn(startPosition, endPosition);
        if (collapsed) {
            getSheet().getSheet().setColumnGroupCollapsed(endPosition, true);
        }
    }

    public <T> T asType(Class<T> type) {
        if (type.isInstance(xssfRow)) {
            return (T)xssfRow;
        }
        return DefaultGroovyMethods.asType(this, type);
    }
}
