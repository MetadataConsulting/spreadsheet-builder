package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.CellMatcher
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition

class PoiCellMatcher implements CellMatcher {

    @Override
    Collection<CellDefinition> match(File spreadsheet, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookDefinition") Closure workbookDefinition) {
        Collection cells = Collections.emptyList()
        spreadsheet.withInputStream {
            cells = match(it, workbookDefinition)
        }
        return cells
    }

    @Override
    Collection<CellDefinition> match(InputStream inputStream, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookDefinition") Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)

        Collection<CellDefinition> cells = collectAllCells(workbook)

        return cells
    }

    private Collection<CellDefinition> collectAllCells(XSSFWorkbook workbook) {
        Collection<CellDefinition> cells = []

        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)

        for (PoiSheetDefinition sheet in poiWorkbook.sheets) {
            for (PoiRowDefinition row in sheet.rows) {
                for (PoiCellDefinition cell in row.cells) {
                    cells << cell
                }
            }
        }
        cells
    }
}
