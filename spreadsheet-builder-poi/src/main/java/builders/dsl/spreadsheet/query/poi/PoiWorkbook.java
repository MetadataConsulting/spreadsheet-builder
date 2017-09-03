package builders.dsl.spreadsheet.query.poi;

import builders.dsl.spreadsheet.api.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

    public List<builders.dsl.spreadsheet.api.Sheet> getSheets() {
        List<builders.dsl.spreadsheet.api.Sheet> sheets = new ArrayList<builders.dsl.spreadsheet.api.Sheet>();
        for (Sheet s : workbook) {
            sheets.add(new PoiSheet(this, (XSSFSheet) s));
        }
        return Collections.unmodifiableList(sheets);
    }

}
