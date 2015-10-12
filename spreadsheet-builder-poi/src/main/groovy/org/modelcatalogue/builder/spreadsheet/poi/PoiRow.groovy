package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.modelcatalogue.builder.spreadsheet.api.Cell
import org.modelcatalogue.builder.spreadsheet.api.CellStyle
import org.modelcatalogue.builder.spreadsheet.api.Row

@CompileStatic class PoiRow implements Row {

    private final XSSFRow xssfRow
    private final PoiSheet sheet

    private String styleName
    private Closure styleDefinition = { return null }

    private final List<Integer> startPositions = []
    private int nextColNumber = 0

    PoiRow(PoiSheet sheet, XSSFRow xssfRow) {
        this.sheet = sheet
        this.xssfRow = xssfRow
    }

    @Override
    void cell() {
        xssfRow.createCell(nextColNumber++, org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK)
    }

    @Override
    void cell(Object value) {
        XSSFCell xssfCell = xssfRow.createCell(nextColNumber++)

        PoiCell poiCell = new PoiCell(this, xssfCell)
        if (styleName) {
            poiCell.style styleName
        }
        poiCell.style styleDefinition
        poiCell.value value
    }

    @Override
    void cell(@DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition) {
        XSSFCell xssfCell = xssfRow.createCell(nextColNumber++)

        PoiCell poiCell = new PoiCell(this, xssfCell)
        if (styleName) {
            poiCell.style styleName
        }
        poiCell.style styleDefinition
        poiCell.with cellDefinition

        handleSpans(xssfCell, poiCell)
    }

    private void handleSpans(XSSFCell xssfCell, PoiCell poiCell) {
        if (poiCell.colspan > 1 || poiCell.rowspan > 1) {
            xssfRow.sheet.addMergedRegion(new CellRangeAddress(
                    xssfCell.rowIndex,
                    xssfCell.rowIndex + poiCell.rowspan - 1,
                    xssfCell.columnIndex,
                    xssfCell.columnIndex + poiCell.colspan - 1
            ));
        }
    }

    @Override
    void cell(int column, @DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition) {
        XSSFCell xssfCell = xssfRow.createCell(column - 1)

        nextColNumber = column

        PoiCell poiCell = new PoiCell(this, xssfCell)
        if (styleName) {
            poiCell.style styleName
        }
        poiCell.style styleDefinition
        poiCell.with cellDefinition

        handleSpans(xssfCell, poiCell)
    }

    @Override
    void cell(String column, @DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition) {
        cell parseColumn(column), cellDefinition
    }

    @Override
    void style(@DelegatesTo(CellStyle.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyle") Closure styleDefinition) {
        this.styleDefinition = styleDefinition
    }

    @Override
    void style(String name) {
        this.styleName = name
    }

    protected PoiSheet getSheet() {
        return sheet
    }

    protected XSSFRow getRow() {
        return xssfRow
    }

    @Override
    void group(@DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure insideGroupDefinition) {
        createGroup(false, insideGroupDefinition)
    }

    @Override
    void collapse(@DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure insideGroupDefinition) {
        createGroup(true, insideGroupDefinition)
    }

    private void createGroup(boolean collapsed, @DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure insideGroupDefinition) {
        startPositions.push nextColNumber
        with insideGroupDefinition

        int startPosition = startPositions.pop()

        if (nextColNumber - startPosition > 1) {
            int endPosition = nextColNumber - 1
            sheet.sheet.groupColumn(startPosition, endPosition)
            if (collapsed) {
                sheet.sheet.setColumnGroupCollapsed(endPosition, true)
            }
        }

    }

    protected static int parseColumn(String column) {
        char a = 'A'
        char[] chars = column.toUpperCase().toCharArray().toList().reverse()
        int acc = 0
        for (int i = chars.size() ; i-- ; i > 0) {
            if (i == 0) {
                acc += ((int) chars[i].charValue() - (int) a.charValue() + 1)
            } else {
                acc += 26 * i * ((int) chars[i].charValue() - (int) a.charValue() + 1)
            }
        }
        return acc
    }
}
