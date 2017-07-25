package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractWorkbookDefinition;

class PoiWorkbookDefinition extends AbstractWorkbookDefinition implements WorkbookDefinition {

    private final XSSFWorkbook workbook;

    PoiWorkbookDefinition(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    @Override
    protected PoiSheetDefinition createSheet(String name) {
        XSSFSheet sheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name));
        return new PoiSheetDefinition(this, sheet != null ? sheet : workbook.createSheet(WorkbookUtil.createSafeSheetName(name)));
    }

    @Override
    protected PoiCellStyleDefinition createCellStyle() {
        return new PoiCellStyleDefinition(this);
    }

    XSSFWorkbook getWorkbook() {
        return workbook;
    }


    <T> T asType(Class<T> type) {
        if (type.isInstance(workbook)) {
            return type.cast(workbook);
        }
        return DefaultGroovyMethods.asType(this, type);
    }

    void addPendingLink(String ref, PoiCellDefinition cell) {
        addPendingLink(new PoiPendingLink(cell, ref));
    }

}
