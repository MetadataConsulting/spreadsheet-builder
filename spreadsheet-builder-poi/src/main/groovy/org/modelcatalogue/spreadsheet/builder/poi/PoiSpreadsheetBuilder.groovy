package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
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

    @Override
    void build(File file, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.WorkbookDefinition") Closure workbookDefinition) {
        file.withOutputStream {
            build it, workbookDefinition
        }
    }
}
