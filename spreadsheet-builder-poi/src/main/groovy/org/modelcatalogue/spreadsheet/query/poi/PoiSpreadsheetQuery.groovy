package org.modelcatalogue.spreadsheet.query.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Workbook
import org.modelcatalogue.spreadsheet.builder.poi.PoiWorkbookDefinition
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetQuery
import org.modelcatalogue.spreadsheet.query.simple.SimpleSpreadsheetQuery
import org.modelcatalogue.spreadsheet.query.simple.WorkbookSupplier

class PoiSpreadsheetQuery {

    static  SpreadsheetQuery create() {
        new SimpleSpreadsheetQuery(new WorkbookSupplier() {
            @Override
            Workbook load(InputStream stream) {
                return new PoiWorkbookDefinition(new XSSFWorkbook(stream))
            }
        })
    }

}
