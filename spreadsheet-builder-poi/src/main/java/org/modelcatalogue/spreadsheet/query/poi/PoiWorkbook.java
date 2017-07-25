package org.modelcatalogue.spreadsheet.query.poi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelcatalogue.spreadsheet.api.Workbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

class PoiWorkbook implements Workbook {

    private final XSSFWorkbook workbook;

    PoiWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public List<org.modelcatalogue.spreadsheet.api.Sheet> getSheets() {
        List<org.modelcatalogue.spreadsheet.api.Sheet> sheets = new ArrayList<org.modelcatalogue.spreadsheet.api.Sheet>();
        for (Sheet s : workbook) {
            sheets.add(new PoiSheet(this, (XSSFSheet) s));
        }
        return Collections.unmodifiableList(sheets);
    }

}
