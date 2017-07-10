package org.modelcatalogue.spreadsheet.builder.poi;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition;
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractRowDefinition;

class PoiRowDefinition extends AbstractRowDefinition implements RowDefinition {

    private final XSSFRow xssfRow;

    PoiRowDefinition(PoiSheetDefinition sheet, XSSFRow xssfRow) {
        super(sheet);
        this.xssfRow = xssfRow;
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

    public PoiSheetDefinition getSheet() {
        return (PoiSheetDefinition) sheet;
    }

    protected XSSFRow getRow() {
        return xssfRow;
    }


    protected void doCreateGroup(int startPosition, int endPosition, boolean collapsed) {
        getSheet().getSheet().groupColumn(startPosition, endPosition);
        if (collapsed) {
            getSheet().getSheet().setColumnGroupCollapsed(endPosition, true);
        }
    }

}
