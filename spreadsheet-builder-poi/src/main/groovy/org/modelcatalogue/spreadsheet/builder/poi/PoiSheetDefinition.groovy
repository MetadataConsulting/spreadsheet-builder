package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.api.Page
import org.modelcatalogue.spreadsheet.api.Sheet
import org.modelcatalogue.spreadsheet.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition

class PoiSheetDefinition implements SheetDefinition, Sheet {

    private final XSSFSheet xssfSheet
    private final PoiWorkbookDefinition workbook

    private final List<Integer> startPositions = []
    private int nextRowNumber = 0
    private final Set<Integer> autoColumns = new HashSet<Integer>()
    private final Map<Integer, PoiRowDefinition> rows = [:]
    private boolean automaticFilter

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
    PoiSheetDefinition row() {
        findOrCreateRow nextRowNumber++
        this
    }

    @Override
    Sheet getNext() {
        int current = workbook.workbook.getSheetIndex(sheet.getSheetName())

        if (current == workbook.workbook.getNumberOfSheets() - 1) {
            return null
        }
        XSSFSheet next = workbook.workbook.getSheetAt(current + 1)
        Sheet nextPoiSheet = workbook.sheets.find { it.sheet.sheetName == next.sheetName }
        if (nextPoiSheet) {
            return nextPoiSheet
        }
        return new PoiSheetDefinition(workbook, next)
    }

    @Override
    Sheet getPrevious() {
        int current = workbook.workbook.getSheetIndex(sheet.getSheetName())

        if (current == 0) {
            return null
        }
        XSSFSheet next = workbook.workbook.getSheetAt(current - 1)
        Sheet nextPoiSheet = workbook.sheets.find { it.sheet.sheetName == next.sheetName }
        if (nextPoiSheet) {
            return nextPoiSheet
        }
        return new PoiSheetDefinition(workbook, next)
    }

    @Override
    Page getPage() {
        return new PoiPageSettingsProvider(this)
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
    PoiSheetDefinition row(Configurer<RowDefinition> rowDefinition) {
        PoiRowDefinition row = findOrCreateRow nextRowNumber++
        Configurer.Runner.doConfigure(rowDefinition, row)
        this
    }

    @Override
    PoiSheetDefinition row(int oneBasedRowNumber, Configurer<RowDefinition> rowDefinition) {
        assert oneBasedRowNumber > 0
        nextRowNumber = oneBasedRowNumber

        PoiRowDefinition poiRow = findOrCreateRow oneBasedRowNumber - 1
        Configurer.Runner.doConfigure(rowDefinition, poiRow)
        this
    }

    @Override
    PoiSheetDefinition freeze(int column, int row) {
        xssfSheet.createFreezePane(column, row)
        this
    }

    @Override
    PoiSheetDefinition freeze(String column, int row) {
        freeze Cell.Util.parseColumn(column), row
        this
    }

    @Override
    PoiSheetDefinition collapse(Configurer<SheetDefinition> insideGroupDefinition) {
        createGroup(true, insideGroupDefinition)
        this
    }

    @Override
    PoiSheetDefinition group(Configurer<SheetDefinition> insideGroupDefinition) {
        createGroup(false, insideGroupDefinition)
        this
    }

    @Override
    PoiSheetDefinition lock() {
        sheet.enableLocking()
        return null
    }

    @Override
    PoiSheetDefinition password(String password) {
        sheet.protectSheet(password)
        this
    }

    @Override
    PoiSheetDefinition filter(Keywords.Auto auto) {
        automaticFilter = true
        this
    }

    @Override
    PoiSheetDefinition page(Configurer<PageDefinition> pageDefinition) {
        PageDefinition page = new PoiPageSettingsProvider(this)
        Configurer.Runner.doConfigure(pageDefinition, page)
        this
    }

    private void createGroup(boolean collapsed, Configurer<SheetDefinition> insideGroupDefinition) {
        startPositions.push nextRowNumber
        Configurer.Runner.doConfigure(insideGroupDefinition, this)

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

    public <T> T asType(Class<T> type) {
        if (type.isInstance(sheet)) {
            return sheet as T
        }
        return super.asType(type)
    }

    protected void processAutomaticFilter() {
        if (automaticFilter && sheet.lastRowNum > 0) {
            sheet.setAutoFilter(new CellRangeAddress(
                    sheet.firstRowNum,
                    sheet.lastRowNum,
                    sheet.getRow(sheet.firstRowNum).firstCellNum,
                    sheet.getRow(sheet.firstRowNum).lastCellNum - 1
            ))
        }
    }
}
