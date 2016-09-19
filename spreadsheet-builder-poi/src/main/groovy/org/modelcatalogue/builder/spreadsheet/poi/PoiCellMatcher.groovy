package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.Cell
import org.modelcatalogue.builder.spreadsheet.api.CellMatcher
import org.modelcatalogue.builder.spreadsheet.api.Workbook

class PoiCellMatcher implements CellMatcher {

    @Override
    Collection<Cell> match(File spreadsheet, @DelegatesTo(Workbook.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Workbook") Closure workbookDefinition) {
        Collection cells = Collections.emptyList()
        spreadsheet.withInputStream {
            cells = match(it, workbookDefinition)
        }
        return cells
    }

    @Override
    Collection<Cell> match(InputStream inputStream, @DelegatesTo(Workbook.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Workbook") Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)

        Collection<Cell> cells = collectAllCells(workbook)

        return cells
    }

    private Collection<Cell> collectAllCells(XSSFWorkbook workbook) {
        Collection<Cell> cells = []

        PoiWorkbook poiWorkbook = new PoiWorkbook(workbook)

        for (PoiSheet sheet in poiWorkbook.sheets) {
            for (PoiRow row in sheet.rows) {
                for (PoiCell cell in row.cells) {
                    cells << cell
                }
            }
        }
        cells
    }
}
