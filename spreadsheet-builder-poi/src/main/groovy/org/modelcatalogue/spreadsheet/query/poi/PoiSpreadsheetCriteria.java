package org.modelcatalogue.spreadsheet.query.poi;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria;
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaFactory;
import org.modelcatalogue.spreadsheet.query.simple.SimpleSpreadsheetCriteria;

import java.io.*;

public enum PoiSpreadsheetCriteria implements SpreadsheetCriteriaFactory {
    FACTORY;

    @Override
    public SpreadsheetCriteria forFile(File spreadsheet) throws FileNotFoundException {
        return forStream(new FileInputStream(spreadsheet));
    }

    @Override
    public SpreadsheetCriteria forStream(InputStream stream) {
        try {
            return SimpleSpreadsheetCriteria.forWorkbook(new PoiWorkbook(new XSSFWorkbook(stream)));
        } catch (IOException e) {
            throw new RuntimeException("Exception creating new workbook: " + stream, e);
        }
    }

}
