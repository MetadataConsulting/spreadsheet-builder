package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.SpreadsheetBuilder
import org.modelcatalogue.builder.spreadsheet.api.Workbook


@CompileStatic class PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    @Override
    void build(OutputStream outputStream, @DelegatesTo(Workbook.class) Closure workbookDefinition) {
        XSSFWorkbook workbook = new XSSFWorkbook()

        PoiWorkbook poiWorkbook = new PoiWorkbook(workbook)
        poiWorkbook.with workbookDefinition

        workbook.write(outputStream)
    }
}
