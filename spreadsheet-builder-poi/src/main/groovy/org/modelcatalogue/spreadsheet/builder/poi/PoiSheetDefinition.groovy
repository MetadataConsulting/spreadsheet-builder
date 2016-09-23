package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.api.Sheet
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition

class PoiSheetDefinition implements SheetDefinition, Sheet {

    private final XSSFSheet xssfSheet
    private final PoiWorkbookDefinition workbook

    private final List<Integer> startPositions = []
    private int nextRowNumber = 0
    private final Set<Integer> autoColumns = new HashSet<Integer>()
    private final Map<Integer, PoiRowDefinition> rows = [:]

    PoiSheetDefinition(PoiWorkbookDefinition workbook, XSSFSheet xssfSheet) {
        this.workbook = workbook
        this.xssfSheet = xssfSheet
    }


    @Override
    String getName() {
        return xssfSheet.sheetName
    }

    @Override
    PoiWorkbookDefinition getWorkbook() {
        return workbook
    }

    List<PoiRowDefinition> getRows() {
        xssfSheet.collect {
            PoiRowDefinition row = getRowByNumber(it.rowNum + 1)
            if (row) {
                return row
            }
            return createRowWrapper(it.rowNum + 1)
        }
    }

    PoiRowDefinition getRowByNumber(int rowNumberStartingOne) {
        rows[rowNumberStartingOne]
    }

    @Override
    void row() {
        findOrCreateRow nextRowNumber++
    }

    @Override
    Sheet getNext() {
        int current = workbook.workbook.getSheetIndex(sheet.getSheetName());

        if (current == workbook.workbook.getNumberOfSheets() - 1) {
            return null
        }
        XSSFSheet next = workbook.workbook.getSheetAt(current + 1);
        Sheet nextPoiSheet = workbook.sheets.find { it.sheet.sheetName == next.sheetName }
        if (nextPoiSheet) {
            return nextPoiSheet
        }
        return new PoiSheetDefinition(workbook, next)
    }

    @Override
    Sheet getPrevious() {
        int current = workbook.workbook.getSheetIndex(sheet.getSheetName());

        if (current == 0) {
            return null
        }
        XSSFSheet next = workbook.workbook.getSheetAt(current - 1);
        Sheet nextPoiSheet = workbook.sheets.find { it.sheet.sheetName == next.sheetName }
        if (nextPoiSheet) {
            return nextPoiSheet
        }
        return new PoiSheetDefinition(workbook, next)
    }

    private PoiRowDefinition findOrCreateRow(int zeroBasedRowNumber) {
        PoiRowDefinition row = rows[zeroBasedRowNumber + 1]

        if (row) {
            return row
        }

        XSSFRow xssfRow = xssfSheet.createRow(zeroBasedRowNumber)

        row = new PoiRowDefinition(this, xssfRow)

        rows[zeroBasedRowNumber + 1] = row

        return row
    }

    @Override
    void row(@DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        PoiRowDefinition row = findOrCreateRow nextRowNumber++
        row.with rowDefinition
    }

    @Override
    void row(int oneBasedRowNumber, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        assert oneBasedRowNumber > 0
        nextRowNumber = oneBasedRowNumber

        PoiRowDefinition poiRow = findOrCreateRow oneBasedRowNumber - 1
        poiRow.with rowDefinition
    }

    @Override
    void freeze(int column, int row) {
        xssfSheet.createFreezePane(column, row)
    }

    @Override
    void freeze(String column, int row) {
        freeze Cell.Util.parseColumn(column), row
    }

    @Override
    void collapse(@DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        createGroup(true, insideGroupDefinition)
    }

    @Override
    void group(@DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        createGroup(false, insideGroupDefinition)
    }

    @Override
    Object getLocked() {
        sheet.enableLocking()
        return null
    }

    @Override
    void password(String password) {
        sheet.protectSheet(password)
    }

    private void createGroup(boolean collapsed, @DelegatesTo(SheetDefinition.class) Closure insideGroupDefinition) {
        startPositions.push nextRowNumber
        with insideGroupDefinition

        int startPosition = startPositions.pop()

        if (nextRowNumber - startPosition > 1) {
            int endPosition = nextRowNumber - 1
            xssfSheet.groupRow(startPosition, endPosition)
            if (collapsed) {
                xssfSheet.setRowGroupCollapsed(endPosition, true)
            }
        }

    }

    protected XSSFSheet getSheet() {
        return xssfSheet
    }

    protected void addAutoColumn(int i) {
        autoColumns << i
    }

    protected void processAutoColumns() {
        for (int index in autoColumns) {
            xssfSheet.autoSizeColumn(index)
        }
    }

    PoiRowDefinition createRowWrapper(int oneBasedRowNumber) {
        rows[oneBasedRowNumber] = new PoiRowDefinition(this, sheet.getRow(oneBasedRowNumber - 1))
    }

    @Override
    String toString() {
        return "Sheet[${name}]"
    }
}
