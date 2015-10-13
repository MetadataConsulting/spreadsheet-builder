package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFName
import org.codehaus.groovy.runtime.StringGroovyMethods
import org.modelcatalogue.builder.spreadsheet.api.AbstractCell
import org.modelcatalogue.builder.spreadsheet.api.AutoKeyword
import org.modelcatalogue.builder.spreadsheet.api.CellStyle
import org.modelcatalogue.builder.spreadsheet.api.Comment
import org.modelcatalogue.builder.spreadsheet.api.LinkDefinition
import org.modelcatalogue.builder.spreadsheet.api.ToKeyword

class PoiCell extends AbstractCell {

    private final PoiRow row
    private final XSSFCell xssfCell

    private int colspan = 1
    private int rowspan = 1

    private PoiCellStyle poiCellStyle

    PoiCell(PoiRow row, XSSFCell xssfCell) {
        this.xssfCell = xssfCell
        this.row = row
    }

    @Override
    void value(Object value) {
        if (!value) {
            xssfCell.setCellType(Cell.CELL_TYPE_BLANK)
            return
        }

        if (value instanceof Number) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value.doubleValue())
            return
        }

        if (value instanceof Date) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value as Date)
            return
        }

        if (value instanceof Calendar) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value as Calendar)
            return
        }

        if (value instanceof Boolean) {
            xssfCell.setCellType(Cell.CELL_TYPE_BOOLEAN)
            xssfCell.setCellValue(value as Boolean)
            return
        }

        if (value instanceof Boolean) {
            xssfCell.setCellType(Cell.CELL_TYPE_BOOLEAN)
            xssfCell.setCellValue(value as Boolean)
            return
        }

        if (value instanceof CharSequence) {
            value = StringGroovyMethods.stripIndent(value as CharSequence).trim()
        }

        xssfCell.setCellType(Cell.CELL_TYPE_STRING)
        xssfCell.setCellValue(value.toString())
    }

    @Override
    void style(@DelegatesTo(CellStyle.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyle") Closure styleDefinition) {
        if (!poiCellStyle) {
            poiCellStyle = new PoiCellStyle(xssfCell)
        }
        poiCellStyle.with styleDefinition
    }

    @Override
    void comment(String commentText) {
        comment {
            text commentText
        }
    }

    @Override
    void formula(String formula) {
        row.sheet.workbook.addPendingFormula(formula, xssfCell)
    }

    @Override
    void comment(@DelegatesTo(Comment.class) Closure commentDefinition) {
        PoiComment poiComment = new PoiComment()
        poiComment.with commentDefinition
        poiComment.applyTo xssfCell
    }

    @Override
    void colspan(int span) {
        this.colspan = span
    }

    @Override
    void rowspan(int span) {
        this.rowspan = span
    }

    @Override
    void style(String name) {
        style row.sheet.workbook.getStyle(name)
    }

    @Override
    void name(String name) {
        XSSFName theName = xssfCell.row.sheet.workbook.createName() as XSSFName
        theName.setNameName(fixName(name))
        theName.setRefersToFormula("'${xssfCell.sheet.sheetName.replaceAll(/'/, /\'/)}'!${xssfCell.reference}")
    }

    protected static String fixName(String name) {
        if (!name) { throw new IllegalArgumentException("Name cannot be null or empty!") }
        if (name in ['c', 'C', 'r', 'R']) {
            return "_$name"
        }
        name = name.replaceAll(/[^\.0-9a-zA-Z_]/, '_')
        if (!(name =~ /^[abd-qs-zABD-QS-Z_]/)) {
            return fixName("_$name")
        }
        return name
    }

    @Override
    LinkDefinition link(ToKeyword to) {
        return new PoiLinkDefinition(row.sheet.workbook, xssfCell)
    }

    @Override
    void width(double width) {
        row.sheet.sheet.setColumnWidth(xssfCell.columnIndex, (int)Math.round(width * 255D))
    }

    @Override
    void height(double height) {
        row.row.setHeightInPoints(height.floatValue());
    }

    @Override
    void width(AutoKeyword auto) {
        row.sheet.addAutoColumn(xssfCell.columnIndex)
    }

    protected int getColspan() {
        return colspan
    }

    protected int getRowspan() {
        return rowspan
    }
}