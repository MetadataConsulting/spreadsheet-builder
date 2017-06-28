package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition


enum PoiSpreadsheetBuilder implements SpreadsheetBuilder {

    INSTANCE

    @Override
    SpreadsheetDefinition build(Configurer<WorkbookDefinition> workbookDefinition) {
        return buildInternal(new XSSFWorkbook(), workbookDefinition)
    }

    private static PoiWorkbookDefinition buildInternal(XSSFWorkbook workbook, Configurer<WorkbookDefinition> workbookDefinition) {
        PoiWorkbookDefinition poiWorkbook = new PoiWorkbookDefinition(workbook)
        Configurer.Runner.doConfigure(workbookDefinition, poiWorkbook)
        poiWorkbook.resolve()

        return poiWorkbook
    }

    @Override
    SpreadsheetDefinition build(InputStream template, Configurer<WorkbookDefinition> workbookDefinition) {
        template.withStream {
            buildInternal(new XSSFWorkbook(it), workbookDefinition)
        }
    }

    @Override
    SpreadsheetDefinition build(File template, Configurer<WorkbookDefinition> workbookDefinition) {
        template.withInputStream {
            buildInternal(new XSSFWorkbook(it), workbookDefinition)
        }
    }
}
