package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition


enum PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    INSTANCE;

    @Override
    SpreadsheetDefinition build( @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        return buildInternal(new XSSFWorkbook(), workbookDefinition)
    }

    private static PoiWorkbookDefinition buildInternal(XSSFWorkbook workbook, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)
        poiWorkbook.with workbookDefinition
        poiWorkbook.resolve()

        return poiWorkbook
    }

    @Override
    SpreadsheetDefinition build(InputStream template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        template.withStream {
            buildInternal(new XSSFWorkbook(it), workbookDefinition)
        }
    }

    @Override
    SpreadsheetDefinition build(File template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        template.withInputStream {
            buildInternal(new XSSFWorkbook(it), workbookDefinition)
        }
    }
}
