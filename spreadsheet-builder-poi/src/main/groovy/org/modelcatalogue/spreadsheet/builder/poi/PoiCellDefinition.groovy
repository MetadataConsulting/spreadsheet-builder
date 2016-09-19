package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Name
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFComment
import org.apache.poi.xssf.usermodel.XSSFName
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.codehaus.groovy.runtime.StringGroovyMethods

import org.modelcatalogue.spreadsheet.api.Cell as SpreadsheetCell
import org.modelcatalogue.spreadsheet.api.Comment
import org.modelcatalogue.spreadsheet.builder.api.AbstractCellDefinition

import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.CommentDefinition
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition
import org.modelcatalogue.spreadsheet.builder.api.ImageCreator

import org.modelcatalogue.spreadsheet.builder.api.Keywords
import org.modelcatalogue.spreadsheet.builder.api.LinkDefinition

class PoiCellDefinition extends AbstractCellDefinition implements Resolvable, SpreadsheetCell {

    private final PoiRowDefinition row
    private final XSSFCell xssfCell

    private int colspan = 1
    private int rowspan = 1

    private PoiCellStyleDefinition poiCellStyle

    private List<RichTextPart> richTextParts = []

    PoiCellDefinition(PoiRowDefinition row, XSSFCell xssfCell) {
        this.xssfCell = xssfCell
        this.row = row
    }

    @Override
    int getColumn() {
        return xssfCell.columnIndex + 1
    }

    @Override
    String getColumnAsString() {
        return SpreadsheetCell.Util.toColumn(getColumn())
    }

    @Override
    def <T> T read(Class<T> type) {
        if (CharSequence.isAssignableFrom(type)) {
            return xssfCell.stringCellValue as T
        }

        if (Date.isAssignableFrom(type)) {
            return xssfCell.dateCellValue as T
        }

        if (Boolean.isAssignableFrom(type)) {
            return xssfCell.booleanCellValue as T
        }

        if (Number.isAssignableFrom(type)) {
            Double val = xssfCell.getNumericCellValue()
            if (val == null) {
                return null
            }
            return val.asType(type)
        }

        throw new IllegalArgumentException("Cannot read value ${xssfCell.rawValue} of cell as $type")
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
    void style(@DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyleDefinition") Closure styleDefinition) {
        if (!poiCellStyle) {
            poiCellStyle = new PoiCellStyleDefinition(this)
        }
        poiCellStyle.checkSealed()
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
    void comment(@DelegatesTo(CommentDefinition.class) Closure commentDefinition) {
        PoiCommentDefinition poiComment = new PoiCommentDefinition()
        poiComment.with commentDefinition
        poiComment.applyTo xssfCell
    }

    @Override
    Comment getComment() {
        XSSFComment comment = xssfCell.getCellComment()
        if (!comment) {
            return new PoiCommentDefinition()
        }
        return new PoiCommentDefinition(author: comment.author, text: comment.string.string)
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
        if (!poiCellStyle) {
            poiCellStyle = row.sheet.workbook.getStyle(name)
            poiCellStyle.assignTo(this)
            return
        }
        poiCellStyle.checkSealed()
        poiCellStyle.with row.sheet.workbook.getStyleDefinition(name)
    }

    @Override
    void styles(String... names) {
        if (!poiCellStyle) {
            poiCellStyle = row.sheet.workbook.getStyles(names)
            poiCellStyle.assignTo(this)
            return
        }
        poiCellStyle.checkSealed()
        for (String name in names) {
            poiCellStyle.with row.sheet.workbook.getStyleDefinition(name)
        }
    }

    @Override
    void styles(Iterable<String> names) {
        styles(names.toList().toArray(new String[names.size()]))
    }

    @Override
    void style(String name, @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyleDefinition") Closure styleDefinition) {
        style row.sheet.workbook.getStyleDefinition(name)
        style styleDefinition
    }

    @Override
    void name(String name) {
        XSSFName theName = xssfCell.row.sheet.workbook.createName() as XSSFName
        theName.setNameName(fixName(name))
        theName.setRefersToFormula(generateRefersToFormula())
    }

    private String generateRefersToFormula() {
        "'${xssfCell.sheet.sheetName.replaceAll(/'/, /\'/)}'!${xssfCell.reference}"
    }

    @Override
    String getName() {
        Workbook wb = xssfCell.sheet.workbook

        new CellReference(xssfCell).formatAsString()
        List<String> possibleReferences = [new CellReference(xssfCell).formatAsString(), generateRefersToFormula()]
        for (int nn=0; nn< wb.getNumberOfNames(); nn++) {
            Name n = wb.getNameAt(nn);
            for (String reference in possibleReferences) {
                if (n.sheetIndex == -1 || n.sheetIndex == wb.getSheetIndex(xssfCell.sheet)) {
                    if (n.refersToFormula == reference) {
                        return n.nameName;
                    }
                }
            }
        }
        return null
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
    LinkDefinition link(Keywords.To to) {
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
    void width(Keywords.Auto auto) {
        row.sheet.addAutoColumn(xssfCell.columnIndex)
    }

    @Override
    void text(String run) {
        text(run, null)
    }

    @Override
    void text(String run, @DelegatesTo(FontDefinition.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.FontDefinition") Closure fontConfiguration) {
        if (!run) {
            return
        }
        int start = 0
        if (richTextParts) {
            start = richTextParts.last().end
        }
        int end = start + run.length()

        if (!fontConfiguration) {
            richTextParts << new RichTextPart(run, null, start, end)
            return
        }

        PoiFontDefinition font = new PoiFontDefinition(xssfCell.row.sheet.workbook)
        font.with fontConfiguration

        richTextParts << new RichTextPart(run, font, start, end)
    }

    @Override
    ImageCreator png(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_PNG)
    }

    @Override
    ImageCreator jpeg(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_JPEG)
    }

    @Override
    ImageCreator pict(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_JPEG)
    }

    @Override
    ImageCreator emf(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_EMF)
    }

    @Override
    ImageCreator wmf(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_WMF)
    }

    @Override
    ImageCreator dib(Keywords.Image image) {
        return createImageConfigurer(Workbook.PICTURE_TYPE_DIB)
    }

    protected ImageCreator createImageConfigurer(int fileType) {
        return new PoiImageCreator(this, fileType)
    }

    protected int getColspan() {
        return colspan
    }

    protected int getRowspan() {
        return rowspan
    }

    protected XSSFCell getCell() {
        return xssfCell
    }

    @Override
    PoiRowDefinition getRow() {
        return row
    }

    @Override
    void resolve() {
        if (richTextParts) {
            XSSFRichTextString text = xssfCell.richStringCellValue

            text.string = richTextParts.collect { it.text }.join('')

            for (RichTextPart richTextPart in richTextParts) {
                if (richTextPart.text && richTextPart.font) {
                    text.applyFont(richTextPart.start, richTextPart.end, richTextPart.font.font)
                }
            }

            xssfCell.setCellValue(text)
        }
        if ((colspan > 1 || rowspan > 1) && poiCellStyle) {
            poiCellStyle.setBorderTo(cellRangeAddress, row.sheet)
        }
    }

    CellRangeAddress getCellRangeAddress() {
        new CellRangeAddress(
                xssfCell.rowIndex,
                xssfCell.rowIndex + rowspan - 1,
                xssfCell.columnIndex,
                xssfCell.columnIndex + colspan - 1
        )
    }
}
