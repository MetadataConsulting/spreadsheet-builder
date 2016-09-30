package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Workbook
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.Stylesheet
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition

class PoiWorkbookDefinition implements WorkbookDefinition, Workbook, SpreadsheetDefinition {

    private final XSSFWorkbook workbook
    private final Map<String, Closure> namedStylesDefinition = [:]
    private final Map<String, PoiCellStyleDefinition> namedStyles = [:]
    private final Map<String, PoiSheetDefinition> sheets = [:]
    private final List<Resolvable> toBeResolved = []

    PoiWorkbookDefinition(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    @Override
    void sheet(String name, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure sheetDefinition) {
        PoiSheetDefinition sheet = sheets[name]

        if (!sheet) {
            XSSFSheet xssfSheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name)) ?: workbook.createSheet(WorkbookUtil.createSafeSheetName(name))
            sheet = new PoiSheetDefinition(this, xssfSheet)
            sheets[name] = sheet
        }

        sheet.with sheetDefinition

        sheet.processAutoColumns()
    }

    @Override
    void style(String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        namedStylesDefinition[name] = styleDefinition
    }

    @Override
    void apply(Class<? extends Stylesheet> stylesheet) {
        apply stylesheet.newInstance()
    }

    @Override
    void apply(Stylesheet stylesheet) {
        stylesheet.declareStyles(this)
    }

    XSSFWorkbook getWorkbook() {
        return workbook
    }

    protected PoiCellStyleDefinition getStyle(String name) {
        PoiCellStyleDefinition style = namedStyles[name]

        if (style) {
            return style
        }

        style = new PoiCellStyleDefinition(this)
        style.with getStyleDefinition(name)
        style.seal()

        namedStyles[name] = style

        return style
    }

    protected PoiCellStyleDefinition getStyles(String... names) {
        String name = names.join('.')

        PoiCellStyleDefinition style = namedStyles[name]

        if (style) {
            return style
        }

        style = new PoiCellStyleDefinition(this)
        for (String n in names) {
            style.with getStyleDefinition(n)
        }
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
}
