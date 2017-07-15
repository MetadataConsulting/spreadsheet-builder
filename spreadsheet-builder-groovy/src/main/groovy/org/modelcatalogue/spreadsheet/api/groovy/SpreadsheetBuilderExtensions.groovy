package org.modelcatalogue.spreadsheet.api.groovy

import groovy.transform.CompileStatic
import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.modelcatalogue.spreadsheet.api.BorderPositionProvider
import org.modelcatalogue.spreadsheet.api.BorderStyle
import org.modelcatalogue.spreadsheet.api.BorderStyleProvider
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.api.DataRow
import org.modelcatalogue.spreadsheet.api.FontStyle
import org.modelcatalogue.spreadsheet.api.FontStylesProvider
import org.modelcatalogue.spreadsheet.api.ForegroundFill
import org.modelcatalogue.spreadsheet.api.ForegroundFillProvider
import org.modelcatalogue.spreadsheet.api.ColorProvider
import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.api.Keywords.VerticalAlignment
import org.modelcatalogue.spreadsheet.api.PageSettingsProvider
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition
import org.modelcatalogue.spreadsheet.builder.api.CanDefineStyle
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.CommentDefinition
import org.modelcatalogue.spreadsheet.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.DimensionModifier
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition
import org.modelcatalogue.spreadsheet.builder.api.HasStyle
import org.modelcatalogue.spreadsheet.builder.api.PageDefinition
import org.modelcatalogue.spreadsheet.builder.api.RowDefinition
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition
import org.modelcatalogue.spreadsheet.query.api.BorderCriterion
import org.modelcatalogue.spreadsheet.query.api.CellCriterion
import org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion
import org.modelcatalogue.spreadsheet.query.api.FontCriterion
import org.modelcatalogue.spreadsheet.query.api.PageCriterion
import org.modelcatalogue.spreadsheet.query.api.RowCriterion
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaResult
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion

/**
 * Main purpose of this class is to provide additional context for IDEs and static type checking.
 */
@CompileStatic class SpreadsheetBuilderExtensions {

    static CellDefinition(CellDefinition self, CharSequence sequence) {
        self.value(sequence.stripIndent().trim())
    }

    static CanDefineStyle style(CanDefineStyle stylable, String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(name, styleDefinition as Configurer<CellStyleDefinition>)
    }

    static CellDefinition comment(CellDefinition cellDefinition, @DelegatesTo(CommentDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CommentDefinition") Closure commentDefinition) {
        cellDefinition.comment(commentDefinition as Configurer<CommentDefinition>)
    }

    static CellDefinition text(CellDefinition cellDefinition, String text, @DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.FontDefinition") Closure fontConfiguration) {
        cellDefinition.text(text, fontConfiguration as Configurer<FontDefinition>)
    }

    static CellStyleDefinition font(CellStyleDefinition style, @DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.FontDefinition") Closure fontConfiguration) {
        style.font(fontConfiguration as Configurer<FontDefinition>)
    }

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    static CellStyleDefinition border(CellStyleDefinition style, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleDefinition border(CellStyleDefinition style, Keywords.BorderSide location, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(location, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleDefinition border(CellStyleDefinition style, Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(first, second, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleDefinition border(CellStyleDefinition style, Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(first, second, third, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Applies a customized named style to the current element.
     *
     * @param name the name of the style
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    static HasStyle style(HasStyle stylable, String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(name, styleDefinition as Configurer<CellStyleDefinition>)
    }

    /**
     * Applies a customized named style to the current element.
     *
     * @param names the names of the styles
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    static HasStyle styles(HasStyle stylable, Iterable<String> names, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.styles(names, styleDefinition as Configurer<CellStyleDefinition>)
    }

    /**
     * Applies the style defined by the closure to the current element.
     * @param styleDefinition the definition of the style
     */
    static HasStyle style(HasStyle stylable, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(styleDefinition as Configurer<CellStyleDefinition>)
    }

    static RowDefinition cell(RowDefinition row, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(cellDefinition as Configurer<CellDefinition>)
    }
    static RowDefinition cell(RowDefinition row, int column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(column, cellDefinition as Configurer<CellDefinition>)
    }
    static RowDefinition cell(RowDefinition row, String column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(column, cellDefinition as Configurer<CellDefinition>)
    }
    static RowDefinition group(RowDefinition row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure insideGroupDefinition) {
        row.group(insideGroupDefinition as Configurer<RowDefinition>)
    }
    static RowDefinition collapse(RowDefinition row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure insideGroupDefinition) {
        row.collapse(insideGroupDefinition as Configurer<RowDefinition>)
    }

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition closure defining the content of the row
     */
    static SheetDefinition row(SheetDefinition sheet, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        sheet.row(rowDefinition as Configurer<RowDefinition>)
    }

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition closure defining the content of the row
     */
    static SheetDefinition row(SheetDefinition sheet, int row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        sheet.row(row, rowDefinition as Configurer<RowDefinition>)
    }

    static SheetDefinition group(SheetDefinition sheet, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        sheet.group(insideGroupDefinition as Configurer<SheetDefinition>)
    }
    static SheetDefinition collapse(SheetDefinition sheet, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        sheet.collapse(insideGroupDefinition as Configurer<SheetDefinition>)
    }

    /**
     * Configures the basic page settings.
     * @param pageDefinition closure defining the page settings
     */
    static SheetDefinition page(SheetDefinition sheet, @DelegatesTo(PageDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.PageDefinition") Closure pageDefinition) {
        sheet.page(pageDefinition as Configurer<PageDefinition>)
    }

    static SheetDefinition lock(SheetDefinition self, SheetDefinition otherSelf) {
        otherSelf.lock()
    }

    static SheetDefinition show(SheetDefinition self, SheetDefinition otherSelf) {
        otherSelf.show()
    }

    static SheetDefinition hide(SheetDefinition self, SheetDefinition otherSelf) {
        otherSelf.hide()
    }

    static SheetDefinition hideCompletely(SheetDefinition self, SheetDefinition otherSelf) {
        otherSelf.hideCompletely()
    }

    static SheetCriterion show(SheetCriterion self, SheetCriterion otherSelf) {
        otherSelf.show()
    }

    static SheetCriterion hide(SheetCriterion self, SheetCriterion otherSelf) {
        otherSelf.hide()
    }

    static SheetCriterion hideCompletely(SheetCriterion self, SheetCriterion otherSelf) {
        otherSelf.hideCompletely()
    }

    static SpreadsheetDefinition build(SpreadsheetBuilder builder, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        builder.build(workbookDefinition as Configurer<WorkbookDefinition>)
    }
    static SpreadsheetDefinition build(SpreadsheetBuilder builder, InputStream template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        builder.build(template, workbookDefinition as Configurer<WorkbookDefinition>)
    }
    static SpreadsheetDefinition build(SpreadsheetBuilder builder, File template, @DelegatesTo(WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition") Closure workbookDefinition) {
        builder.build(template, workbookDefinition as Configurer<WorkbookDefinition>)
    }

    static WorkbookDefinition sheet(WorkbookDefinition workbook, String name, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure sheetDefinition) {
        workbook.sheet(name, sheetDefinition as Configurer<SheetDefinition>)
    }

    static CellCriterion style(CellCriterion cell, @DelegatesTo(CellStyleCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion") Closure styleCriterion) {
        cell.style(styleCriterion as Configurer<CellStyleCriterion>)
    }
    static CellCriterion or(CellCriterion cell, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure sheetCriterion) {
        cell.or(sheetCriterion as Configurer<CellCriterion>)
    }


    static CellStyleCriterion font(CellStyleCriterion style, @DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion) {
        style.font(fontCriterion as Configurer<FontCriterion>)
    }

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    static CellStyleCriterion border(CellStyleCriterion style, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleCriterion border(CellStyleCriterion style, Keywords.BorderSide location, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(location, borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleCriterion border(CellStyleCriterion style, Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(first, second, borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    static CellStyleCriterion border(CellStyleCriterion style, Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(first, second, third, borderConfiguration as Configurer<BorderCriterion>)
    }

    static RowCriterion cell(RowCriterion row, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(cellCriterion as Configurer<CellCriterion>)
    }
    static RowCriterion cell(RowCriterion row, int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(column, cellCriterion as Configurer<CellCriterion>)
    }
    static RowCriterion cell(RowCriterion row, String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(column, cellCriterion as Configurer<CellCriterion>)
    }
    static RowCriterion cell(RowCriterion row, int from, int to, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(from, to, cellCriterion as Configurer<CellCriterion>)
    }
    static RowCriterion cell(RowCriterion row, String from, String to, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(from, to, cellCriterion as Configurer<CellCriterion>)
    }

    static RowCriterion or(RowCriterion row, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        row.or(rowCriterion as Configurer<RowCriterion>)
    }

    static SheetCriterion row(SheetCriterion sheet, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        sheet.row(rowCriterion as Configurer<RowCriterion>)
    }
    static SheetCriterion row(SheetCriterion sheet, int row, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        sheet.row(row, rowCriterion as Configurer<RowCriterion>)
    }
    static SheetCriterion page(SheetCriterion sheet, @DelegatesTo(PageCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.PageCriterion") Closure pageCriterion) {
        sheet.page(pageCriterion as Configurer<PageCriterion>)
    }
    static SheetCriterion or(SheetCriterion sheet, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        sheet.or(sheetCriterion as Configurer<SheetCriterion>)
    }

    static SpreadsheetCriteriaResult query(SpreadsheetCriteria criteria, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        criteria.query(workbookCriterion as Configurer<WorkbookCriterion>)
    }
    static Cell find(SpreadsheetCriteria criteria, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        criteria.find(workbookCriterion as Configurer<WorkbookCriterion>)
    }
    static boolean exists(SpreadsheetCriteria criteria, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) throws FileNotFoundException {
        criteria.exists(workbookCriterion as Configurer<WorkbookCriterion>)
    }

    static WorkbookCriterion sheet(WorkbookCriterion workbook, String name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        workbook.sheet(name, sheetCriterion as Configurer<SheetCriterion>)
    }
    static WorkbookCriterion sheet(WorkbookCriterion workbook, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        workbook.sheet(sheetCriterion as Configurer<SheetCriterion>)
    }
    static WorkbookCriterion or(WorkbookCriterion workbook, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        workbook.or(workbookCriterion as Configurer<WorkbookCriterion>)
    }

    static Cell getAt(DataRow self, String name) {
        self.get(name)
    }

    /**
     * Converts the dimension to centimeters.
     *
     * This feature is currently experimental.
     */
    static CellDefinition getCm(DimensionModifier self){
        self.cm()
    }

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    static CellDefinition getInch(DimensionModifier self){
        self.inch()
    }

    /**
     * Converts the dimension to inches.
     *
     * This feature is currently experimental.
     */
    static CellDefinition getInches(DimensionModifier self){
        self.inches()
    }

    /**
     * Keeps the dimesion in points.
     *
     * This feature is currently experimental.
     */
    static CellDefinition getPoints(DimensionModifier self){
        self.points()
    }

    static HorizontalAlignmentConfigurer align(CellStyleDefinition self, VerticalAlignment verticalAlignment) {
        new DefaultHorizontalAlignmentConfigurer(self, verticalAlignment);
    }

    static BorderStyle getNone(BorderStyleProvider self) {
        return BorderStyle.NONE
    }

    static BorderStyle getThin(BorderStyleProvider self) {
        return BorderStyle.THIN
    }

    static BorderStyle getMedium(BorderStyleProvider self) {
        return BorderStyle.MEDIUM
    }

    static BorderStyle getDashed(BorderStyleProvider self) {
        return BorderStyle.DASHED
    }

    static BorderStyle getDotted(BorderStyleProvider self) {
        return BorderStyle.DOTTED
    }

    static BorderStyle getThick(BorderStyleProvider self) {
        return BorderStyle.THICK
    }

    static BorderStyle getDouble(BorderStyleProvider self) {
        return BorderStyle.DOUBLE
    }

    static BorderStyle getHair(BorderStyleProvider self) {
        return BorderStyle.HAIR
    }

    static BorderStyle getMediumDashed(BorderStyleProvider self) {
        return BorderStyle.MEDIUM_DASHED
    }

    static BorderStyle getDashDot(BorderStyleProvider self) {
        return BorderStyle.DASH_DOT
    }

    static BorderStyle getMediumDashDot(BorderStyleProvider self) {
        return BorderStyle.MEDIUM_DASH_DOT
    }

    static BorderStyle getDashDotDot(BorderStyleProvider self) {
        return BorderStyle.DASH_DOT_DOT
    }

    static BorderStyle getMediumDashDotDot(BorderStyleProvider self) {
        return BorderStyle.MEDIUM_DASH_DOT_DOT
    }

    static BorderStyle getSlantedDashDot(BorderStyleProvider self) {
        return BorderStyle.SLANTED_DASH_DOT
    }

    static Keywords.Auto getAuto(CellDefinition self) {
        return Keywords.Auto.AUTO
    }

    static Keywords.To getTo(CellDefinition self) {
        return Keywords.To.TO
    }

    static Keywords.Image getImage(CellDefinition self) {
        return Keywords.Image.IMAGE
    }

    static ForegroundFill getNoFill(ForegroundFillProvider self) {
        return ForegroundFill.NO_FILL
    }

    static ForegroundFill getSolidForeground(ForegroundFillProvider self) {
        return ForegroundFill.SOLID_FOREGROUND
    }

    static ForegroundFill getFineDots(ForegroundFillProvider self) {
        return ForegroundFill.FINE_DOTS
    }

    static ForegroundFill getAltBars(ForegroundFillProvider self) {
        return ForegroundFill.ALT_BARS
    }

    static ForegroundFill getSparseDots(ForegroundFillProvider self) {
        return ForegroundFill.SPARSE_DOTS
    }

    static ForegroundFill getThickHorizontalBands(ForegroundFillProvider self) {
        return ForegroundFill.THICK_HORZ_BANDS
    }

    static ForegroundFill getThickVerticalBands(ForegroundFillProvider self) {
        return ForegroundFill.THICK_VERT_BANDS
    }

    static ForegroundFill getThickBackwardDiagonals(ForegroundFillProvider self) {
        return ForegroundFill.THICK_BACKWARD_DIAG
    }

    static ForegroundFill getThickForwardDiagonals(ForegroundFillProvider self) {
        return ForegroundFill.THICK_FORWARD_DIAG
    }

    static ForegroundFill getBigSpots(ForegroundFillProvider self) {
        return ForegroundFill.BIG_SPOTS
    }

    static ForegroundFill getBricks(ForegroundFillProvider self) {
        return ForegroundFill.BRICKS
    }

    static ForegroundFill getThinHorizontalBands(ForegroundFillProvider self) {
        return ForegroundFill.THIN_HORZ_BANDS
    }

    static ForegroundFill getThinVerticalBands(ForegroundFillProvider self) {
        return ForegroundFill.THIN_VERT_BANDS
    }

    static ForegroundFill getThinBackwardDiagonals(ForegroundFillProvider self) {
        return ForegroundFill.THIN_BACKWARD_DIAG
    }

    static ForegroundFill getThinForwardDiagonals(ForegroundFillProvider self) {
        return ForegroundFill.THICK_FORWARD_DIAG
    }

    static ForegroundFill getSquares(ForegroundFillProvider self) {
        return ForegroundFill.SQUARES
    }

    static ForegroundFill getDiamonds(ForegroundFillProvider self) {
        return ForegroundFill.DIAMONDS
    }

    static Keywords.BorderSideAndHorizontalAlignment getLeft(CellStyleDefinition self) {
        return Keywords.BorderSideAndHorizontalAlignment.LEFT
    }

    static Keywords.BorderSideAndHorizontalAlignment getRight(CellStyleDefinition self) {
        return Keywords.BorderSideAndHorizontalAlignment.RIGHT
    }

    static Keywords.BorderSideAndVerticalAlignment getTop(CellStyleDefinition self) {
        return Keywords.BorderSideAndVerticalAlignment.TOP
    }

    static Keywords.BorderSideAndVerticalAlignment getBottom(CellStyleDefinition self) {
        return Keywords.BorderSideAndVerticalAlignment.BOTTOM
    }

    static Keywords.VerticalAndHorizontalAlignment getCenter(CellStyleDefinition self) {
        return Keywords.VerticalAndHorizontalAlignment.CENTER
    }

    static Keywords.VerticalAndHorizontalAlignment getJustify(CellStyleDefinition self) {
        return Keywords.VerticalAndHorizontalAlignment.JUSTIFY
    }

    static Keywords.PureVerticalAlignment getDistributed(CellStyleDefinition self) {
        return Keywords.PureVerticalAlignment.DISTRIBUTED
    }

    static Keywords.Text getText(CellStyleDefinition self) {
        return Keywords.Text.WRAP
    }

    static Keywords.Orientation getPortrait(PageSettingsProvider self) {
        return Keywords.Orientation.PORTRAIT
    }

    static Keywords.Orientation getLandscape(PageSettingsProvider self) {
        return Keywords.Orientation.LANDSCAPE
    }

    static Keywords.Fit getWidth(PageSettingsProvider self) {
        return Keywords.Fit.WIDTH
    }

    static Keywords.Fit getHeight(PageSettingsProvider self) {
        return Keywords.Fit.HEIGHT
    }

    static Keywords.To getTo(PageSettingsProvider self) {
        return Keywords.To.TO
    }

    static Keywords.Paper getLetter(PageSettingsProvider self) {
        return Keywords.Paper.LETTER
    }

    static Keywords.Paper getLetterSmall(PageSettingsProvider self) {
        return Keywords.Paper.LETTER_SMALL
    }

    static Keywords.Paper getTabloid(PageSettingsProvider self) {
        return Keywords.Paper.TABLOID
    }

    static Keywords.Paper getLedger(PageSettingsProvider self) {
        return Keywords.Paper.LEDGER
    }

    static Keywords.Paper getLegal(PageSettingsProvider self) {
        return Keywords.Paper.LEGAL
    }

    static Keywords.Paper getStatement(PageSettingsProvider self) {
        return Keywords.Paper.STATEMENT
    }

    static Keywords.Paper getExecutive(PageSettingsProvider self) {
        return Keywords.Paper.EXECUTIVE
    }

    static Keywords.Paper getA3(PageSettingsProvider self) {
        return Keywords.Paper.A3
    }

    static Keywords.Paper getA4(PageSettingsProvider self) {
        return Keywords.Paper.A4
    }

    static Keywords.Paper getA4Small(PageSettingsProvider self) {
        return Keywords.Paper.A4_SMALL
    }

    static Keywords.Paper getA5(PageSettingsProvider self) {
        return Keywords.Paper.A5
    }

    static Keywords.Paper getB4(PageSettingsProvider self) {
        return Keywords.Paper.B4
    }

    static Keywords.Paper getB5(PageSettingsProvider self) {
        return Keywords.Paper.B5
    }

    static Keywords.Paper getFolio(PageSettingsProvider self) {
        return Keywords.Paper.FOLIO
    }

    static Keywords.Paper getQuarto(PageSettingsProvider self) {
        return Keywords.Paper.QUARTO
    }

    static Keywords.Paper getStandard10x14(PageSettingsProvider self) {
        return Keywords.Paper.STANDARD_10_14
    }

    static Keywords.Paper getStandard11x17(PageSettingsProvider self) {
        return Keywords.Paper.STANDARD_11_17
    }

    static FontStyle getItalic(FontStylesProvider self) {
        FontStyle.ITALIC
    }

    static FontStyle getBold(FontStylesProvider self) {
        FontStyle.BOLD
    }

    static FontStyle getStrikeout(FontStylesProvider self) {
        FontStyle.STRIKEOUT
    }

    static FontStyle getUnderline(FontStylesProvider self) {
        FontStyle.UNDERLINE
    }

    static Keywords.BorderSideAndHorizontalAlignment getLeft(BorderPositionProvider self) {
        return Keywords.BorderSideAndHorizontalAlignment.LEFT;
    }

    static Keywords.BorderSideAndHorizontalAlignment getRight(BorderPositionProvider self) {
        return Keywords.BorderSideAndHorizontalAlignment.RIGHT;
    }

    static Keywords.BorderSideAndVerticalAlignment getTop(BorderPositionProvider self) {
        return Keywords.BorderSideAndVerticalAlignment.TOP;
    }

    static Keywords.BorderSideAndVerticalAlignment getBottom(BorderPositionProvider self) {
        return Keywords.BorderSideAndVerticalAlignment.BOTTOM;
    }

    static Keywords.PureHorizontalAlignment getGeneral(CellStyleDefinition self) {
        return Keywords.PureHorizontalAlignment.GENERAL
    }
    static Keywords.PureHorizontalAlignment getFill(CellStyleDefinition self) {
        return Keywords.PureHorizontalAlignment.FILL
    }
    static Keywords.PureHorizontalAlignment getCenterSelection(CellStyleDefinition self) {
        return Keywords.PureHorizontalAlignment.CENTER_SELECTION
    }
    static Keywords.Auto getAuto(SheetDefinition self) {
        return Keywords.Auto.AUTO
    }

    static Color getAliceBlue(ColorProvider colorProvider) { return Color.aliceBlue }
    static Color getAntiqueWhite(ColorProvider colorProvider) { return Color.antiqueWhite }
    static Color getAqua(ColorProvider colorProvider) { return Color.aqua }
    static Color getAquamarine(ColorProvider colorProvider) { return Color.aquamarine }
    static Color getAzure(ColorProvider colorProvider) { return Color.azure }
    static Color getBeige(ColorProvider colorProvider) { return Color.beige }
    static Color getBisque(ColorProvider colorProvider) { return Color.bisque }
    static Color getBlack(ColorProvider colorProvider) { return Color.black }
    static Color getBlanchedAlmond(ColorProvider colorProvider) { return Color.blanchedAlmond }
    static Color getBlue(ColorProvider colorProvider) { return Color.blue }
    static Color getBlueViolet(ColorProvider colorProvider) { return Color.blueViolet }
    static Color getBrown(ColorProvider colorProvider) { return Color.brown }
    static Color getBurlyWood(ColorProvider colorProvider) { return Color.burlyWood }
    static Color getCadetBlue(ColorProvider colorProvider) { return Color.cadetBlue }
    static Color getChartreuse(ColorProvider colorProvider) { return Color.chartreuse }
    static Color getChocolate(ColorProvider colorProvider) { return Color.chocolate }
    static Color getCoral(ColorProvider colorProvider) { return Color.coral }
    static Color getCornflowerBlue(ColorProvider colorProvider) { return Color.cornflowerBlue }
    static Color getCornsilk(ColorProvider colorProvider) { return Color.cornsilk }
    static Color getCrimson(ColorProvider colorProvider) { return Color.crimson }
    static Color getCyan(ColorProvider colorProvider) { return Color.cyan }
    static Color getDarkBlue(ColorProvider colorProvider) { return Color.darkBlue }
    static Color getDarkCyan(ColorProvider colorProvider) { return Color.darkCyan }
    static Color getDarkGoldenRod(ColorProvider colorProvider) { return Color.darkGoldenRod }
    static Color getDarkGray(ColorProvider colorProvider) { return Color.darkGray }
    static Color getDarkGreen(ColorProvider colorProvider) { return Color.darkGreen }
    static Color getDarkKhaki(ColorProvider colorProvider) { return Color.darkKhaki }
    static Color getDarkMagenta(ColorProvider colorProvider) { return Color.darkMagenta }
    static Color getDarkOliveGreen(ColorProvider colorProvider) { return Color.darkOliveGreen }
    static Color getDarkOrange(ColorProvider colorProvider) { return Color.darkOrange }
    static Color getDarkOrchid(ColorProvider colorProvider) { return Color.darkOrchid }
    static Color getDarkRed(ColorProvider colorProvider) { return Color.darkRed }
    static Color getDarkSalmon(ColorProvider colorProvider) { return Color.darkSalmon }
    static Color getDarkSeaGreen(ColorProvider colorProvider) { return Color.darkSeaGreen }
    static Color getDarkSlateBlue(ColorProvider colorProvider) { return Color.darkSlateBlue }
    static Color getDarkSlateGray(ColorProvider colorProvider) { return Color.darkSlateGray }
    static Color getDarkTurquoise(ColorProvider colorProvider) { return Color.darkTurquoise }
    static Color getDarkViolet(ColorProvider colorProvider) { return Color.darkViolet }
    static Color getDeepPink(ColorProvider colorProvider) { return Color.deepPink }
    static Color getDeepSkyBlue(ColorProvider colorProvider) { return Color.deepSkyBlue }
    static Color getDimGray(ColorProvider colorProvider) { return Color.dimGray }
    static Color getDodgerBlue(ColorProvider colorProvider) { return Color.dodgerBlue }
    static Color getFireBrick(ColorProvider colorProvider) { return Color.fireBrick }
    static Color getFloralWhite(ColorProvider colorProvider) { return Color.floralWhite }
    static Color getForestGreen(ColorProvider colorProvider) { return Color.forestGreen }
    static Color getFuchsia(ColorProvider colorProvider) { return Color.fuchsia }
    static Color getGainsboro(ColorProvider colorProvider) { return Color.gainsboro }
    static Color getGhostWhite(ColorProvider colorProvider) { return Color.ghostWhite }
    static Color getGold(ColorProvider colorProvider) { return Color.gold }
    static Color getGoldenRod(ColorProvider colorProvider) { return Color.goldenRod }
    static Color getGray(ColorProvider colorProvider) { return Color.gray }
    static Color getGreen(ColorProvider colorProvider) { return Color.green }
    static Color getGreenYellow(ColorProvider colorProvider) { return Color.greenYellow }
    static Color getHoneyDew(ColorProvider colorProvider) { return Color.honeyDew }
    static Color getHotPink(ColorProvider colorProvider) { return Color.hotPink }
    static Color getIndianRed (ColorProvider colorProvider) { return Color.indianRed  }
    static Color getIndigo (ColorProvider colorProvider) { return Color.indigo  }
    static Color getIvory(ColorProvider colorProvider) { return Color.ivory }
    static Color getKhaki(ColorProvider colorProvider) { return Color.khaki }
    static Color getLavender(ColorProvider colorProvider) { return Color.lavender }
    static Color getLavenderBlush(ColorProvider colorProvider) { return Color.lavenderBlush }
    static Color getLawnGreen(ColorProvider colorProvider) { return Color.lawnGreen }
    static Color getLemonChiffon(ColorProvider colorProvider) { return Color.lemonChiffon }
    static Color getLightBlue(ColorProvider colorProvider) { return Color.lightBlue }
    static Color getLightCoral(ColorProvider colorProvider) { return Color.lightCoral }
    static Color getLightCyan(ColorProvider colorProvider) { return Color.lightCyan }
    static Color getLightGoldenRodYellow(ColorProvider colorProvider) { return Color.lightGoldenRodYellow }
    static Color getLightGray(ColorProvider colorProvider) { return Color.lightGray }
    static Color getLightGreen(ColorProvider colorProvider) { return Color.lightGreen }
    static Color getLightPink(ColorProvider colorProvider) { return Color.lightPink }
    static Color getLightSalmon(ColorProvider colorProvider) { return Color.lightSalmon }
    static Color getLightSeaGreen(ColorProvider colorProvider) { return Color.lightSeaGreen }
    static Color getLightSkyBlue(ColorProvider colorProvider) { return Color.lightSkyBlue }
    static Color getLightSlateGray(ColorProvider colorProvider) { return Color.lightSlateGray }
    static Color getLightSteelBlue(ColorProvider colorProvider) { return Color.lightSteelBlue }
    static Color getLightYellow(ColorProvider colorProvider) { return Color.lightYellow }
    static Color getLime(ColorProvider colorProvider) { return Color.lime }
    static Color getLimeGreen(ColorProvider colorProvider) { return Color.limeGreen }
    static Color getLinen(ColorProvider colorProvider) { return Color.linen }
    static Color getMagenta(ColorProvider colorProvider) { return Color.magenta }
    static Color getMaroon(ColorProvider colorProvider) { return Color.maroon }
    static Color getMediumAquaMarine(ColorProvider colorProvider) { return Color.mediumAquaMarine }
    static Color getMediumBlue(ColorProvider colorProvider) { return Color.mediumBlue }
    static Color getMediumOrchid(ColorProvider colorProvider) { return Color.mediumOrchid }
    static Color getMediumPurple(ColorProvider colorProvider) { return Color.mediumPurple }
    static Color getMediumSeaGreen(ColorProvider colorProvider) { return Color.mediumSeaGreen }
    static Color getMediumSlateBlue(ColorProvider colorProvider) { return Color.mediumSlateBlue }
    static Color getMediumSpringGreen(ColorProvider colorProvider) { return Color.mediumSpringGreen }
    static Color getMediumTurquoise(ColorProvider colorProvider) { return Color.mediumTurquoise }
    static Color getMediumVioletRed(ColorProvider colorProvider) { return Color.mediumVioletRed }
    static Color getMidnightBlue(ColorProvider colorProvider) { return Color.midnightBlue }
    static Color getMintCream(ColorProvider colorProvider) { return Color.mintCream }
    static Color getMistyRose(ColorProvider colorProvider) { return Color.mistyRose }
    static Color getMoccasin(ColorProvider colorProvider) { return Color.moccasin }
    static Color getNavajoWhite(ColorProvider colorProvider) { return Color.navajoWhite }
    static Color getNavy(ColorProvider colorProvider) { return Color.navy }
    static Color getOldLace(ColorProvider colorProvider) { return Color.oldLace }
    static Color getOlive(ColorProvider colorProvider) { return Color.olive }
    static Color getOliveDrab(ColorProvider colorProvider) { return Color.oliveDrab }
    static Color getOrange(ColorProvider colorProvider) { return Color.orange }
    static Color getOrangeRed(ColorProvider colorProvider) { return Color.orangeRed }
    static Color getOrchid(ColorProvider colorProvider) { return Color.orchid }
    static Color getPaleGoldenRod(ColorProvider colorProvider) { return Color.paleGoldenRod }
    static Color getPaleGreen(ColorProvider colorProvider) { return Color.paleGreen }
    static Color getPaleTurquoise(ColorProvider colorProvider) { return Color.paleTurquoise }
    static Color getPaleVioletRed(ColorProvider colorProvider) { return Color.paleVioletRed }
    static Color getPapayaWhip(ColorProvider colorProvider) { return Color.papayaWhip }
    static Color getPeachPuff(ColorProvider colorProvider) { return Color.peachPuff }
    static Color getPeru(ColorProvider colorProvider) { return Color.peru }
    static Color getPink(ColorProvider colorProvider) { return Color.pink }
    static Color getPlum(ColorProvider colorProvider) { return Color.plum }
    static Color getPowderBlue(ColorProvider colorProvider) { return Color.powderBlue }
    static Color getPurple(ColorProvider colorProvider) { return Color.purple }
    static Color getRebeccaPurple(ColorProvider colorProvider) { return Color.rebeccaPurple }
    static Color getRed(ColorProvider colorProvider) { return Color.red }
    static Color getRosyBrown(ColorProvider colorProvider) { return Color.rosyBrown }
    static Color getRoyalBlue(ColorProvider colorProvider) { return Color.royalBlue }
    static Color getSaddleBrown(ColorProvider colorProvider) { return Color.saddleBrown }
    static Color getSalmon(ColorProvider colorProvider) { return Color.salmon }
    static Color getSandyBrown(ColorProvider colorProvider) { return Color.sandyBrown }
    static Color getSeaGreen(ColorProvider colorProvider) { return Color.seaGreen }
    static Color getSeaShell(ColorProvider colorProvider) { return Color.seaShell }
    static Color getSienna(ColorProvider colorProvider) { return Color.sienna }
    static Color getSilver(ColorProvider colorProvider) { return Color.silver }
    static Color getSkyBlue(ColorProvider colorProvider) { return Color.skyBlue }
    static Color getSlateBlue(ColorProvider colorProvider) { return Color.slateBlue }
    static Color getSlateGray(ColorProvider colorProvider) { return Color.slateGray }
    static Color getSnow(ColorProvider colorProvider) { return Color.snow }
    static Color getSpringGreen(ColorProvider colorProvider) { return Color.springGreen }
    static Color getSteelBlue(ColorProvider colorProvider) { return Color.steelBlue }
    static Color getTan(ColorProvider colorProvider) { return Color.tan }
    static Color getTeal(ColorProvider colorProvider) { return Color.teal }
    static Color getThistle(ColorProvider colorProvider) { return Color.thistle }
    static Color getTomato(ColorProvider colorProvider) { return Color.tomato }
    static Color getTurquoise(ColorProvider colorProvider) { return Color.turquoise }
    static Color getViolet(ColorProvider colorProvider) { return Color.violet }
    static Color getWheat(ColorProvider colorProvider) { return Color.wheat }
    static Color getWhite(ColorProvider colorProvider) { return Color.white }
    static Color getWhiteSmoke(ColorProvider colorProvider) { return Color.whiteSmoke }
    static Color getYellow(ColorProvider colorProvider) { return Color.yellow }
    static Color getYellowGreen(ColorProvider colorProvider) { return Color.yellowGreen }
}
