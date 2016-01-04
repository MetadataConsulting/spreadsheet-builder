package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.CellStyle
import org.modelcatalogue.builder.spreadsheet.api.Sheet
import org.modelcatalogue.builder.spreadsheet.api.Stylesheet
import org.modelcatalogue.builder.spreadsheet.api.Workbook

class PoiWorkbook implements Workbook {

    private final XSSFWorkbook workbook
    private final Map<String, Closure> namedStylesDefinition = [:]
    private final Map<String, PoiCellStyle> namedStyles = [:]
    private final Map<String, XSSFDataFormat> formats = [:]
    private final List<Resolvable> toBeResolved = []

    PoiWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    @Override
    void sheet(String name, @DelegatesTo(Sheet.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Sheet") Closure sheetDefinition) {
        XSSFSheet xssfSheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name)) ?: workbook.createSheet(WorkbookUtil.createSafeSheetName(name))

        PoiSheet sheet = new PoiSheet(this, xssfSheet)
        sheet.with sheetDefinition

        sheet.processAutoColumns()
    }

    @Override
    void style(String name, @DelegatesTo(CellStyle.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyle") Closure styleDefinition) {
        namedStylesDefinition[name] = styleDefinition
    }

    @Override
    void apply(Class<Stylesheet> stylesheet) {
        apply stylesheet.newInstance()
    }

    @Override
    void apply(Stylesheet stylesheet) {
        stylesheet.declareStyles(this)
    }

    XSSFWorkbook getWorkbook() {
        return workbook
    }

    protected PoiCellStyle getStyle(String name) {
        PoiCellStyle style = namedStyles[name]

        if (style) {
            return style
        }

        style = new PoiCellStyle(this)
        style.with getStyleDefinition(name)
        style.seal()

        namedStyles[name] = style

        return style
    }

    protected Closure getStyleDefinition(String name) {
        Closure style = namedStylesDefinition[name]
        if (!style) {
            throw new IllegalArgumentException("Style '$name' not defined")
        }
        return style
    }

    protected void addPendingFormula(String formula, XSSFCell cell) {
        toBeResolved << new PendingFormula(cell, formula)
    }

    protected void addPendingLink(String ref, XSSFCell cell) {
        toBeResolved << new PendingLink(cell, ref)
    }

    protected void resolve() {
        for (Resolvable resolvable in toBeResolved) {
            resolvable.resolve()
        }
    }
}
