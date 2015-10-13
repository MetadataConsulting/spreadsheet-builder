package org.modelcatalogue.builder.spreadsheet.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.SpreadsheetBuilder
import org.modelcatalogue.builder.spreadsheet.api.Workbook


class PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    @Override
    void build(OutputStream outputStream, @DelegatesTo(Workbook.class) Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook()

        PoiWorkbook poiWorkbook = new PoiWorkbook(workbook)
        poiWorkbook.with workbookDefinition
        poiWorkbook.resolve()

        workbook.write(outputStream)
    }
}
