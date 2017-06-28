package org.modelcatalogue.spreadsheet.builder.poi

import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.tck.AbstractBuilderSpec
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaFactory
import org.modelcatalogue.spreadsheet.query.poi.PoiSpreadsheetCriteria

class PoiExcelBuilderSpec extends AbstractBuilderSpec {

    @Override
    protected SpreadsheetCriteriaFactory createCriteriaFactory() {
        return PoiSpreadsheetCriteria.FACTORY
    }

    @Override
    protected SpreadsheetBuilder createSpreadsheetBuilder() {
        return PoiSpreadsheetBuilder.INSTANCE
    }

}
