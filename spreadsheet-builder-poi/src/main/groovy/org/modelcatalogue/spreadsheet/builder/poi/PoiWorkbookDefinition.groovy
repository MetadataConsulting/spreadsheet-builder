package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.spreadsheet.api.Workbook
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.Stylesheet
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition

class PoiWorkbookDefinition implements WorkbookDefinition, Workbook, SpreadsheetDefinition {

    private final XSSFWorkbook workbook
    private final Map<String, Configurer<CellStyleDefinition>> namedStylesDefinition = [:]
    private final Map<String, PoiCellStyleDefinition> namedStyles = [:]
    private final Map<String, PoiSheetDefinition> sheets = [:]
    private final List<Resolvable> toBeResolved = []

    PoiWorkbookDefinition(XSSFWorkbook workbook) {
        this.workbook = workbook
    }

    @Override
    void sheet(String name, Configurer<SheetDefinition> sheetDefinition) {
        PoiSheetDefinition sheet = sheets[name]

        if (!sheet) {
            XSSFSheet xssfSheet = workbook.getSheet(WorkbookUtil.createSafeSheetName(name)) ?: workbook.createSheet(WorkbookUtil.createSafeSheetName(name))
            sheet = new PoiSheetDefinition(this, xssfSheet)
            sheets[name] = sheet
        }

        sheetDefinition.configure sheet

        sheet.processAutoColumns()
        sheet.processAutomaticFilter()
    }

    @Override
    void style(String name, Configurer<CellStyleDefinition> styleDefinition) {
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
        getStyleDefinition(name)?.configure(style)
        style.seal()

        namedStyles[name] = style

        return style
    }

    protected PoiCellStyleDefinition getStyles(Iterable<String> names) {
        String name = names.join('.')

        PoiCellStyleDefinition style = namedStyles[name]

        if (style) {
            return style
        }

        style = new PoiCellStyleDefinition(this)
        for (String n in names) {
            getStyleDefinition(n)?.configure(style)
        }
        style.seal()

        namedStyles[name] = style

        return style
    }

    protected Configurer<CellStyleDefinition> getStyleDefinition(String name) {
        Configurer<CellStyleDefinition> style = namedStylesDefinition[name]
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

    public <T> T asType(Class<T> type) {
        if (type.isInstance(workbook)) {
            return workbook as T
        }
        return super.asType(type)
    }


}
