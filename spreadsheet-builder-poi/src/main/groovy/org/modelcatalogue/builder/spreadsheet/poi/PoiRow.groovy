package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.modelcatalogue.builder.spreadsheet.api.Cell
import org.modelcatalogue.builder.spreadsheet.api.CellStyle
import org.modelcatalogue.builder.spreadsheet.api.Row

class PoiRow implements Row {

    private final XSSFRow xssfRow
    private final PoiSheet sheet

    private String styleName
    private String[] styleNames
    private Closure styleDefinition

    private final List<Integer> startPositions = []
    private int nextColNumber = 0
    private final Map<Integer, PoiCell> cells = [:]

    PoiRow(PoiSheet sheet, XSSFRow xssfRow) {
        this.sheet = sheet
        this.xssfRow = xssfRow
    }

    private PoiCell findOrCreateCell(int zeroBasedCellNumber) {
        PoiCell cell = cells[zeroBasedCellNumber + 1]

        if (cell) {
            return cell
        }

        XSSFCell xssfCell = xssfRow.createCell(zeroBasedCellNumber)

        cell = new PoiCell(this, xssfCell)

        cells[zeroBasedCellNumber + 1] = cell

        return cell
    }

    @Override
    void cell() {
        cell null
    }

    @Override
    void cell(Object value) {
        PoiCell poiCell = findOrCreateCell nextColNumber++

        if (styleNames) {
            poiCell.styles styleNames
        }
        if (styleName) {
            if (styleDefinition) {
                poiCell.style styleName, styleDefinition
            } else {
                poiCell.style styleName
            }
        } else if(styleDefinition) {
            poiCell.style styleDefinition
        }

        poiCell.value value

        poiCell.resolve()
    }

    @Override
    void cell(@DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition) {
        PoiCell poiCell = findOrCreateCell nextColNumber

        if (styleName) {
            if (styleDefinition) {
                poiCell.style styleName, styleDefinition
            } else {
                poiCell.style styleName
            }
        } else if (styleDefinition) {
            poiCell.style styleDefinition
        }

        poiCell.with cellDefinition

        nextColNumber += poiCell.colspan

        handleSpans(poiCell)

        poiCell.resolve()
    }

    private void handleSpans(PoiCell poiCell) {
        if (poiCell.colspan > 1 || poiCell.rowspan > 1) {
            xssfRow.sheet.addMergedRegion(poiCell.cellRangeAddress);
        }
    }

    @Override
    void cell(int column, @DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition) {
        nextColNumber = column

        PoiCell poiCell = findOrCreateCell column - 1

        if (styleName) {
            if (styleDefinition) {
                poiCell.style styleName, styleDefinition
            } else {
                poiCell.style styleName
            }
        } else if(styleDefinition) {
            poiCell.style styleDefinition
        }

        poiCell.with cellDefinition

        handleSpans(poiCell)

        poiCell.resolve()
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

    @Override
    void style(String name, @DelegatesTo(CellStyle.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyle") Closure styleDefinition) {
        style name
        style styleDefinition
    }

    @Override
    void styles(String... names) {
        this.styleNames = names
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
