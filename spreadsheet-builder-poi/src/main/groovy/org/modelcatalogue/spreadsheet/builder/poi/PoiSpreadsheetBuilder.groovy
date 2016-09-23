package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition


enum PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    INSTANCE;

    @Override
    SpreadsheetDefinition build(@DelegatesTo(WorkbookDefinition.class) Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook()

        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)
        poiWorkbook.with workbookDefinition
        poiWorkbook.resolve()

        return poiWorkbook
    }
}
