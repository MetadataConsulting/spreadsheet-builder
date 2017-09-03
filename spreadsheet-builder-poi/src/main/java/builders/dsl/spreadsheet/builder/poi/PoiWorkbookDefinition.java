package builders.dsl.spreadsheet.builder.poi;

import builders.dsl.spreadsheet.builder.api.WorkbookDefinition;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import builders.dsl.spreadsheet.impl.AbstractWorkbookDefinition;

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

    void addPendingLink(String ref, PoiCellDefinition cell) {
        addPendingLink(new PoiPendingLink(cell, ref));
    }

}
