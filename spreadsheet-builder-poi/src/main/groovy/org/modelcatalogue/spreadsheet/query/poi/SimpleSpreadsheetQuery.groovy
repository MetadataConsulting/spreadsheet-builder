package org.modelcatalogue.spreadsheet.query.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.api.Row
import org.modelcatalogue.spreadsheet.api.Sheet
import org.modelcatalogue.spreadsheet.api.Workbook
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetQuery
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleCellCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleRowCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleSheetCriterion
import org.modelcatalogue.spreadsheet.query.simple.SimpleWorkbookCriterion
import org.modelcatalogue.spreadsheet.query.simple.WorkbookSupplier

public final class SimpleSpreadsheetQuery implements SpreadsheetQuery {

    private final WorkbookSupplier loader;

    SimpleSpreadsheetQuery(WorkbookSupplier loader) {
        this.loader = loader
    }

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
        Workbook poiWorkbook = loader.load(inputStream)

        Collection<Cell> cells = []
        SimpleWorkbookCriterion criterion = new SimpleWorkbookCriterion()
        criterion.with workbookCriterion

        for (Sheet sheet in poiWorkbook.sheets) {
            if (criterion.passesAnyCondition(sheet)) {
                for (Row row in sheet.rows) {
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
