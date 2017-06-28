package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Workbook
import org.modelcatalogue.spreadsheet.impl.AbstractWorkbookDefinition
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition

class PoiWorkbookDefinition extends AbstractWorkbookDefinition<PoiCellStyleDefinition, PoiSheetDefinition, PoiWorkbookDefinition> implements WorkbookDefinition, Workbook, SpreadsheetDefinition {

    private final XSSFWorkbook workbook

    PoiWorkbookDefinition(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    @Override
    protected PoiSheetDefinition createSheet(String name) {
        return new PoiSheetDefinition(this, workbook.getSheet(WorkbookUtil.createSafeSheetName(name)) ?: workbook.createSheet(WorkbookUtil.createSafeSheetName(name)))
    }

    @Override
    protected PoiCellStyleDefinition createCellStyle() {
        return new PoiCellStyleDefinition(this)
    }

    XSSFWorkbook getWorkbook() {
        return workbook
    }

    List<PoiSheetDefinition> getSheets() {
        // TODO: reuse existing sheets
        workbook.collect { new PoiSheetDefinition(this, it as XSSFSheet) }
    }

    @Override
    void writeTo(OutputStream outputStream) {
        workbook.write(outputStream)
    }

    @Override
    void writeTo(File file) {
        file.withOutputStream {
            writeTo(it)
        }
    }

    public <T> T asType(Class<T> type) {
        if (type.isInstance(workbook)) {
            return workbook as T
        }
        return super.asType(type)
    }

}
