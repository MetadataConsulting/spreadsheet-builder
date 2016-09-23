package org.modelcatalogue.spreadsheet.query.poi

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.builder.poi.PoiWorkbookDefinition
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaFactory
import org.modelcatalogue.spreadsheet.query.simple.SimpleSpreadsheetCriteria

enum PoiSpreadsheetCriteria implements SpreadsheetCriteriaFactory {

    FACTORY;

    @Override
    SpreadsheetCriteria forFile(File spreadsheet) throws FileNotFoundException {
        return forStream(new FileInputStream(spreadsheet));
    }

    @Override
    SpreadsheetCriteria forStream(InputStream stream) {
        return SimpleSpreadsheetCriteria.forWorkbook(new PoiWorkbookDefinition(new XSSFWorkbook(stream)))
    }

}
