package org.modelcatalogue.spreadsheet.query.poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.api.Sheet;

import java.util.*;

class PoiSheet implements Sheet {

    private final XSSFSheet xssfSheet;
    private final PoiWorkbook workbook;

    private Map<Integer, PoiRow> rows;

    PoiSheet(PoiWorkbook workbook, XSSFSheet xssfSheet) {
        this.workbook = workbook;
        this.xssfSheet = xssfSheet;
    }

    @Override
    public String getName() {
        return xssfSheet.getSheetName();
    }

    @Override
    public PoiWorkbook getWorkbook() {
        return workbook;
    }

    public List<org.modelcatalogue.spreadsheet.api.Row> getRows() {
        if (rows == null) {
            rows = new LinkedHashMap<Integer, PoiRow>(xssfSheet.getLastRowNum());
            for (Row it : xssfSheet) {
                int oneBasedIndex = it.getRowNum() + 1;
                rows.put(oneBasedIndex, createRowWrapper(oneBasedIndex));
            }
        }

        return Collections.unmodifiableList(new ArrayList<org.modelcatalogue.spreadsheet.api.Row>(rows.values()));
    }

    PoiRow getRowByNumber(int rowNumberStartingOne) {
        if (this.rows == null) {
            this.getRows();
        }
        return this.rows.get(rowNumberStartingOne);
    }

    @Override
    public Sheet getNext() {
        int current = getWorkbook().getWorkbook().getSheetIndex(xssfSheet.getSheetName());

        if (current == getWorkbook().getWorkbook().getNumberOfSheets() - 1) {
            return null;
        }
        XSSFSheet next = getWorkbook().getWorkbook().getSheetAt(current + 1);

        for (Sheet sheet : getWorkbook().getSheets()) {
            if (sheet.getName().equals(next.getSheetName())) {
                return sheet;
            }
        }


        return new PoiSheet(getWorkbook(), next);
    }

    @Override
    public Sheet getPrevious() {
        int current = getWorkbook().getWorkbook().getSheetIndex(xssfSheet.getSheetName());

        if (current == 0) {
            return null;
        }
        XSSFSheet next = getWorkbook().getWorkbook().getSheetAt(current - 1);
        for (Sheet sheet : getWorkbook().getSheets()) {
            if (sheet.getName().equals(next.getSheetName())) {
                return sheet;
            }
        }
        return new PoiSheet(getWorkbook(), next);
    }

    @Override
    public Page getPage() {
        return new PoiPage(this);
    }

    protected XSSFSheet getSheet() {
        return xssfSheet;
    }

    PoiRow createRowWrapper(int oneBasedRowNumber) {
        return new PoiRow(this, xssfSheet.getRow(oneBasedRowNumber - 1));
    }

    @Override
    public boolean isLocked() {
        return xssfSheet.isSheetLocked();
    }

    public boolean isHidden() {
        return xssfSheet.getWorkbook().getSheetVisibility(xssfSheet.getWorkbook().getSheetIndex(xssfSheet)) == SheetVisibility.HIDDEN;
    }

    public boolean isVisible() {
        return xssfSheet.getWorkbook().getSheetVisibility(xssfSheet.getWorkbook().getSheetIndex(xssfSheet)) == SheetVisibility.VISIBLE;
    }

    public boolean isVeryHidden() {
        return xssfSheet.getWorkbook().getSheetVisibility(xssfSheet.getWorkbook().getSheetIndex(xssfSheet)) == SheetVisibility.VERY_HIDDEN;
    }

}
