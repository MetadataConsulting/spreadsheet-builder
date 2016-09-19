package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition


class PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    @Override
    void build(OutputStream outputStream, @DelegatesTo(WorkbookDefinition.class) Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook()

        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)
        poiWorkbook.with workbookDefinition
        poiWorkbook.resolve()

        workbook.write(outputStream)
    }
}
