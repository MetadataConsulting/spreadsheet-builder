package org.modelcatalogue.spreadsheet.api.groovy

import groovy.transform.CompileStatic
import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.modelcatalogue.spreadsheet.api.Cell
import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.api.Row
import org.modelcatalogue.spreadsheet.api.Sheet
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition
import org.modelcatalogue.spreadsheet.builder.api.CanDefineStyle
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.CommentDefinition
import org.modelcatalogue.spreadsheet.builder.api.Configurer
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
import org.modelcatalogue.spreadsheet.query.api.Predicate
import org.modelcatalogue.spreadsheet.query.api.RowCriterion
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteria
import org.modelcatalogue.spreadsheet.query.api.SpreadsheetCriteriaResult
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion

@CompileStatic class SpreadsheetBuilderExtensions {

    static void style(CanDefineStyle stylable, String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(name, styleDefinition as Configurer<CellStyleDefinition>)
    }

    static void comment(CellDefinition cellDefinition, @DelegatesTo(CommentDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CommentDefinition") Closure commentDefinition) {
        cellDefinition.comment(commentDefinition as Configurer<CommentDefinition>)
    }

    static void text(CellDefinition cellDefinition, String text, @DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.FontDefinition") Closure fontConfiguration) {
        cellDefinition.text(text, fontConfiguration as Configurer<FontDefinition>)
    }

    static void font(CellStyleDefinition style, @DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.FontDefinition") Closure fontConfiguration) {
        style.font(fontConfiguration as Configurer<FontDefinition>)
    }

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleDefinition style, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleDefinition style, Keywords.BorderSide location, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(location, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleDefinition style, Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(first, second, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleDefinition style, Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.BorderDefinition") Closure borderConfiguration) {
        style.border(first, second, third, borderConfiguration as Configurer<BorderDefinition>)
    }

    /**
     * Applies a customized named style to the current element.
     *
     * @param name the name of the style
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    static void style(HasStyle stylable, String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(name, styleDefinition as Configurer<CellStyleDefinition>)
    }

    /**
     * Applies a customized named style to the current element.
     *
     * @param names the names of the styles
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    static void styles(HasStyle stylable, Iterable<String> names, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.styles(names, styleDefinition as Configurer<CellStyleDefinition>)
    }

    /**
     * Applies the style defined by the closure to the current element.
     * @param styleDefinition the definition of the style
     */
    static void style(HasStyle stylable, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition) {
        stylable.style(styleDefinition as Configurer<CellStyleDefinition>)
    }

    static void cell(RowDefinition row, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(cellDefinition as Configurer<CellDefinition>)
    }
    static void cell(RowDefinition row, int column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(column, cellDefinition as Configurer<CellDefinition>)
    }
    static void cell(RowDefinition row, String column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellDefinition") Closure cellDefinition) {
        row.cell(column, cellDefinition as Configurer<CellDefinition>)
    }
    static void group(RowDefinition row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure insideGroupDefinition) {
        row.group(insideGroupDefinition as Configurer<RowDefinition>)
    }
    static void collapse(RowDefinition row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure insideGroupDefinition) {
        row.collapse(insideGroupDefinition as Configurer<RowDefinition>)
    }

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition closure defining the content of the row
     */
    static void row(SheetDefinition sheet, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        sheet.row(rowDefinition as Configurer<RowDefinition>)
    }

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition closure defining the content of the row
     */
    static void row(SheetDefinition sheet, int row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition) {
        sheet.row(row, rowDefinition as Configurer<RowDefinition>)
    }

    static void group(SheetDefinition sheet, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        sheet.group(insideGroupDefinition as Configurer<SheetDefinition>)
    }
    static void collapse(SheetDefinition sheet, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition) {
        sheet.collapse(insideGroupDefinition as Configurer<SheetDefinition>)
    }

    /**
     * Configures the basic page settings.
     * @param pageDefinition closure defining the page settings
     */
    static void page(SheetDefinition sheet, @DelegatesTo(PageDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.PageDefinition") Closure pageDefinition) {
        sheet.page(pageDefinition as Configurer<PageDefinition>)
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

    static void sheet(WorkbookDefinition workbook, String name, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure sheetDefinition) {
        workbook.sheet(name, sheetDefinition as Configurer<SheetDefinition>)
    }

    static void style(CellCriterion cell, @DelegatesTo(CellStyleCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion") Closure styleCriterion) {
        cell.style(styleCriterion as Configurer<CellStyleCriterion>)
    }
    static void or(CellCriterion cell, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure sheetCriterion) {
        cell.or(sheetCriterion as Configurer<CellCriterion>)
    }


    static void font(CellStyleCriterion style, @DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion) {
        style.font(fontCriterion as Configurer<FontCriterion>)
    }

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleCriterion style, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleCriterion style, Keywords.BorderSide location, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(location, borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleCriterion style, Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(first, second, borderConfiguration as Configurer<BorderCriterion>)
    }

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    static void border(CellStyleCriterion style, Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration) {
        style.border(first, second, third, borderConfiguration as Configurer<BorderCriterion>)
    }

    static void cell(RowCriterion row, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(cellCriterion as Configurer<CellCriterion>)
    }
    static void cell(RowCriterion row, int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(column, cellCriterion as Configurer<CellCriterion>)
    }
    static void cell(RowCriterion row, String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(column, cellCriterion as Configurer<CellCriterion>)
    }
    static void cell(RowCriterion row, Predicate<Cell> predicate, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        row.cell(predicate, cellCriterion as Configurer<CellCriterion>)
    }
    static void or(RowCriterion row, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        row.or(rowCriterion as Configurer<RowCriterion>)
    }

    static void row(SheetCriterion sheet, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        sheet.row(rowCriterion as Configurer<RowCriterion>)
    }
    static void row(SheetCriterion sheet, int row, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        sheet.row(row, rowCriterion as Configurer<RowCriterion>)
    }
    static void row(SheetCriterion sheet, Predicate<Row> predicate, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.RowCriterion") Closure rowCriterion) {
        sheet.row(predicate, rowCriterion as Configurer<RowCriterion>)
    }
    static void page(SheetCriterion sheet, @DelegatesTo(PageCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.PageCriterion") Closure pageCriterion) {
        sheet.page(pageCriterion as Configurer<PageCriterion>)
    }
    static void or(SheetCriterion sheet, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
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

    static void sheet(WorkbookCriterion workbook, String name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        workbook.sheet(name, sheetCriterion as Configurer<SheetCriterion>)
    }
    static void sheet(WorkbookCriterion workbook, Predicate<Sheet> predicate, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        workbook.sheet(predicate, sheetCriterion as Configurer<SheetCriterion>)
    }
    static void sheet(WorkbookCriterion workbook, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        workbook.sheet(sheetCriterion as Configurer<SheetCriterion>)
    }
    static void or(WorkbookCriterion workbook, @DelegatesTo(WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion") Closure workbookCriterion) {
        workbook.or(workbookCriterion as Configurer<WorkbookCriterion>)
    }

    
}
