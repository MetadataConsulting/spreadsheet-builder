package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.CellStyle
import org.modelcatalogue.builder.spreadsheet.api.Sheet
import org.modelcatalogue.builder.spreadsheet.api.Workbook


@CompileStatic class PoiWorkbook implements Workbook {

    private final XSSFWorkbook workbook
    private final Map<String, Closure> namedStyles = [:]

    PoiWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    @Override
    void sheet(String name, @DelegatesTo(Sheet.class) Closure sheetDefinition) {
        XSSFSheet xssfSheet = workbook.getSheet(name) ?: workbook.createSheet(name)

        PoiSheet sheet = new PoiSheet(this, xssfSheet)
        sheet.with sheetDefinition

        sheet.processAutoColumns()
    }

    @Override
    void style(String name, @DelegatesTo(CellStyle.class) Closure styleDefinition) {
        namedStyles[name] = styleDefinition
    }

    protected Closure getStyle(String name) {
        Closure style = namedStyles[name]
        if (!style) {
            throw new IllegalArgumentException("Style '$name' not defined")
        }
        return style
    }
}
