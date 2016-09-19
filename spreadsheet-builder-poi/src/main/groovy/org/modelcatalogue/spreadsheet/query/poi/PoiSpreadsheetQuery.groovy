package org.modelcatalogue.spreadsheet.query.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.builder.poi.PoiRowDefinition
import org.modelcatalogue.spreadsheet.builder.poi.PoiSheetDefinition
import org.modelcatalogue.spreadsheet.builder.poi.PoiWorkbookDefinition
import org.modelcatalogue.spreadsheet.query.simple.SimpleCellCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleRowCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleSheetCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleWorkbookCriterion
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetQuery
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion

class PoiSpreadsheetQuery implements SpreadsheetQuery {

    @Override
    Collection<Cell> query(File spreadsheet) {
        return query(spreadsheet) { }
    }

    @Override
    Collection<Cell> query(InputStream inputStream) {
        return query(inputStream) { }
    }

    @Override
    Collection<Cell> query(File spreadsheet, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookCriterion") Closure workbookCriterion) {
        Collection cells = Collections.emptyList()
        spreadsheet.withInputStream {
            cells = query(it, workbookCriterion)
        }
        return cells
    }

    @Override
    Collection<Cell> query(InputStream inputStream, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookCriterion") Closure workbookCriterion) {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)

        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion()

        criterion.with workbookCriterion

        Collection<Cell> cells = []

        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)

        for (PoiSheetDefinition sheet in poiWorkbook.sheets) {
            if (criterion.passesAnyCondition(sheet)) {
                for (PoiRowDefinition row in sheet.rows) {
                    if (criterion.criteria) {
                        for (SimpleSheetCriterion sheetCriterion in criterion.criteria) {
                            if (sheetCriterion.passesAnyCondition(row)) {
                                if (sheetCriterion.criteria) {
                                    for (Cell cell in row.cells) {
                                        for (SimpleRowCriterion rowCriterion in sheetCriterion.criteria) {
                                            if (rowCriterion.passesAnyCondition(cell)) {
                                                if (rowCriterion.criteria) {
                                                    for (SimpleCellCriterion cellCriterion in rowCriterion.criteria) {
                                                        if (cellCriterion.passesAnyCondition(cell)) {
                                                            cells << cell
                                                        }
                                                    }
                                                } else {
                                                    cells << cell
                                                }
                                            }
                                        }
                                    }
                                } else {
                                    cells.addAll(row.cells)
                                }
                            }
                        }
                    } else {
                        cells.addAll(row.cells)
                    }
                }
            }
        }

        return cells
    }
}
