package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Name
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFComment
import org.apache.poi.xssf.usermodel.XSSFName
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFRow
import org.codehaus.groovy.runtime.StringGroovyMethods

import org.modelcatalogue.spreadsheet.api.Cell as SpreadsheetCell
import org.modelcatalogue.spreadsheet.api.CellStyle
import org.modelcatalogue.spreadsheet.api.Comment
import org.modelcatalogue.spreadsheet.builder.api.CellDefinition
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.CommentDefinition
import org.modelcatalogue.spreadsheet.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.DimensionModifier
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition
import org.modelcatalogue.spreadsheet.builder.api.ImageCreator

import org.modelcatalogue.spreadsheet.api.Keywords
import org.modelcatalogue.spreadsheet.builder.api.LinkDefinition
import org.modelcatalogue.spreadsheet.builder.api.Resolvable

class PoiCellDefinition implements CellDefinition, Resolvable, SpreadsheetCell {

    private final PoiRowDefinition row
    private final XSSFCell xssfCell

    private int colspan = 0
    private int rowspan = 0

    private PoiCellStyleDefinition poiCellStyle

    private List<RichTextPart> richTextParts = []

    PoiCellDefinition(PoiRowDefinition row, XSSFCell xssfCell) {
        this.xssfCell = checkNotNull(xssfCell, 'Cell')
        this.row = checkNotNull(row, 'Row')
    }

    private static <T> T checkNotNull(T o, String what) {
        if (o == null) {
            throw new IllegalArgumentException("$what cannot be null")
        }
        return o
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
    <T> T read(Class<T> type) {
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
    Object getValue() {
        switch (xssfCell.cellType) {
            case Cell.CELL_TYPE_BLANK: return ''
            case Cell.CELL_TYPE_BOOLEAN: return xssfCell.getBooleanCellValue()
            case Cell.CELL_TYPE_ERROR: return xssfCell.getErrorCellString()
            case Cell.CELL_TYPE_FORMULA: return xssfCell.getCellFormula()
            case Cell.CELL_TYPE_NUMERIC: return xssfCell.getNumericCellValue()
            case Cell.CELL_TYPE_STRING: return xssfCell.getStringCellValue()
        }
        return xssfCell.getRawValue()
    }

    @Override
    PoiCellDefinition value(Object value) {
        if (value == null) {
            xssfCell.setCellType(Cell.CELL_TYPE_BLANK)
            return this
        }

        if (value instanceof Number) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value.doubleValue())
            return this
        }

        if (value instanceof Date) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value as Date)
            return this
        }

        if (value instanceof Calendar) {
            xssfCell.setCellType(Cell.CELL_TYPE_NUMERIC)
            xssfCell.setCellValue(value as Calendar)
            return this
        }

        if (value instanceof Boolean) {
            xssfCell.setCellType(Cell.CELL_TYPE_BOOLEAN)
            xssfCell.setCellValue(value as Boolean)
            return this
        }

        if (value instanceof CharSequence) {
            value = StringGroovyMethods.stripIndent(value as CharSequence).trim()
        }

        xssfCell.setCellType(Cell.CELL_TYPE_STRING)
        xssfCell.setCellValue(value.toString())
        return this
    }

    @Override
    PoiCellDefinition style(Configurer<CellStyleDefinition> styleDefinition) {
        if (!poiCellStyle) {
            poiCellStyle = new PoiCellStyleDefinition(this)
        }
        poiCellStyle.checkSealed()
        Configurer.Runner.doConfigure(styleDefinition, poiCellStyle)
        this
    }

    @Override
    PoiCellDefinition comment(String commentText) {
        comment {
            text commentText
        }
        this
    }

    @Override
    PoiCellDefinition formula(String formula) {
        row.sheet.workbook.addPendingFormula(formula, xssfCell)
        this
    }

    @Override
    PoiCellDefinition comment(Configurer<CommentDefinition> commentDefinition) {
        PoiCommentDefinition poiComment = new PoiCommentDefinition()
        Configurer.Runner.doConfigure(commentDefinition, poiComment)
        poiComment.applyTo xssfCell
        this
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
    CellStyle getStyle() {
        return xssfCell.cellStyle ? new PoiCellStyle(xssfCell.cellStyle) : null
    }

    @Override
    PoiCellDefinition colspan(int span) {
        this.colspan = span
        this
    }

    @Override
    PoiCellDefinition rowspan(int span) {
        this.rowspan = span
        this
    }

    @Override
    PoiCellDefinition style(String name) {
        styles name
        this
    }

    @Override
    PoiCellDefinition styles(String... names) {
        styles(names.toList())
        this
    }

    @Override
    PoiCellDefinition styles(Iterable<String> names) {
        if (!poiCellStyle) {
            poiCellStyle = row.sheet.workbook.getStyles(new LinkedHashSet<String>((names + row.styles).toList()))
            poiCellStyle.assignTo(this)
            return this
        }
        if (poiCellStyle.sealed && row.styles) {
            poiCellStyle = null
            styles(new LinkedHashSet<String>(names.toList() + row.styles))
            return this
        }
        poiCellStyle.checkSealed()
        for (String name in names) {
            Configurer.Runner.doConfigure(row.sheet.workbook.getStyleDefinition(name), poiCellStyle)
        }
        this
    }

    @Override
    PoiCellDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition) {
        style row.sheet.workbook.getStyleDefinition(name)
        style styleDefinition
        this
    }

    @Override
    PoiCellDefinition styles(Iterable<String> names, Configurer<CellStyleDefinition> styleDefinition) {
        for (String name in names) {
            Configurer.Runner.doConfigure(row.sheet.workbook.getStyleDefinition(name), poiCellStyle)
        }
        style styleDefinition
        this
    }

    @Override
    PoiCellDefinition styles(Iterable<String> names, Iterable<Configurer<CellStyleDefinition>> styleDefinition) {
        if (!styleDefinition) {
            if (!names) {
                return this
            }
            if (!poiCellStyle) {
                poiCellStyle = row.sheet.workbook.getStyles(new LinkedHashSet<String>((names + row.styles).toList()))
                poiCellStyle.assignTo(this)
                return this
            }
            if (poiCellStyle.sealed && row.styles) {
                poiCellStyle = null
                styles(new LinkedHashSet<String>(names.toList() + row.styles))
                return this
            }
        }
        if (!poiCellStyle) {
            poiCellStyle = new PoiCellStyleDefinition(this)
        }
        poiCellStyle.checkSealed()
        for (String name in names) {
            Configurer.Runner.doConfigure(row.sheet.workbook.getStyleDefinition(name), poiCellStyle)
        }
        for (Configurer<CellStyleDefinition> configurer in styleDefinition) {
            Configurer.Runner.doConfigure(configurer, poiCellStyle)
        }
        this
    }

    @Override
    PoiCellDefinition name(String name) {
        if (fixName(name) != name) {
            throw new IllegalArgumentException("Name ${name} is not valid Excel name! Suggestion: ${fixName(name)}")
        }
        XSSFName theName = xssfCell.row.sheet.workbook.createName() as XSSFName
        theName.setNameName(name)
        theName.setRefersToFormula(generateRefersToFormula())
        this
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
            Name n = wb.getNameAt(nn)
            for (String reference in possibleReferences) {
                if (n.sheetIndex == -1 || n.sheetIndex == wb.getSheetIndex(xssfCell.sheet)) {
                    if (n.refersToFormula == reference) {
                        return n.nameName
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
        return new PoiLinkDefinition(row.sheet.workbook, this)
    }

    @Override
    DimensionModifier width(double width) {
        row.sheet.sheet.setColumnWidth(xssfCell.columnIndex, (int)Math.round(width * 255D))
        return new PoiWidthModifier(this, width)
    }

    @Override
    DimensionModifier height(double height) {
        row.row.setHeightInPoints(height.floatValue())
        return new PoiHeightModifier(this, height)
    }

    @Override
    PoiCellDefinition width(Keywords.Auto auto) {
        row.sheet.addAutoColumn(xssfCell.columnIndex)
        this
    }

    @Override
    PoiCellDefinition text(String run) {
        text(run, null)
        this
    }

    @Override
    PoiCellDefinition text(String run, Configurer<FontDefinition> fontConfiguration) {
        if (!run) {
            return this
        }
        int start = 0
        if (richTextParts) {
            start = richTextParts.last().end
        }
        int end = start + run.length()

        if (!fontConfiguration) {
            richTextParts << new RichTextPart(run, null, start, end)
            return this
        }

        PoiFontDefinition font = new PoiFontDefinition(xssfCell.row.sheet.workbook)
        Configurer.Runner.doConfigure(fontConfiguration, font)

        richTextParts << new RichTextPart(run, font, start, end)
        this
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

    int getColspan() {
        if (colspan >= 1) {
            return colspan
        }



        if (row.sheet.sheet.numMergedRegions == 0) {
            return colspan = 1
        }

        CellRangeAddress address = row.sheet.sheet.mergedRegions.find {
            it.isInRange(cell.rowIndex, cell.columnIndex)
        }

        if (address) {
            rowspan = address.lastRow - address.firstRow + 1
            colspan = address.lastColumn - address.firstColumn + 1
            return colspan
        }
        return colspan = 1
    }

    int getRowspan() {
        if (rowspan >= 1) {
            return rowspan
        }

        if (row.sheet.sheet.numMergedRegions == 0) {
            return rowspan = 1
        }

        CellRangeAddress address = row.sheet.sheet.mergedRegions.find {
            it.isInRange(cell.rowIndex, cell.columnIndex)
        }

        if (address) {
            rowspan = address.lastRow - address.firstRow + 1
            colspan = address.lastColumn - address.firstColumn + 1
            return rowspan
        }
        return rowspan = 1
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
            // XXX: setting border messes up the style of the cell
            poiCellStyle.assignTo(this)
        }
    }

    CellRangeAddress getCellRangeAddress() {
        new CellRangeAddress(
                xssfCell.rowIndex,
                xssfCell.rowIndex + getRowspan() - 1,
                xssfCell.columnIndex,
                xssfCell.columnIndex + getColspan() - 1
        )
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getAbove() {
        PoiRowDefinition row = this.row.getAbove()
        if (!row) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column - 1))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getBellow() {
        PoiRowDefinition row = this.row.getBellow(getRowspan())
        if (!row) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column - 1))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getLeft() {
        if (column == 1) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column - 1)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column - 2))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getRight() {
        if (column + getColspan() > row.row.lastCellNum) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column + getColspan())

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column + getColspan() - 1))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getAboveLeft() {
        PoiRowDefinition row = this.row.getAbove()
        if (!row) {
            return null
        }
        if (column == 1) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column - 1)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column - 2))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getAboveRight() {
        PoiRowDefinition row = this.row.getAbove()
        if (!row) {
            return null
        }
        if (column == row.row.lastCellNum) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column + 1)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getBellowLeft() {
        PoiRowDefinition row = this.row.getBellow()
        if (!row) {
            return null
        }
        if (column == 1) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column - 1)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column - 2))
    }

    @Override
    org.modelcatalogue.spreadsheet.api.Cell getBellowRight() {
        PoiRowDefinition row = this.row.getBellow()
        if (!row) {
            return null
        }
        if (column == row.row.lastCellNum) {
            return null
        }
        PoiCellDefinition existing = row.getCellByNumber(column + 1)

        if (existing) {
            return existing
        }

        return createCellIfExists(getCell(row.row, column))
    }


    protected static XSSFCell getCell(XSSFRow row, int column) {
        XSSFCell cell = row.getCell(column)
        if (cell) {
            return cell
        }
        if (row.sheet.numMergedRegions == 0) {
            return null
        }
        CellRangeAddress address = row.sheet.mergedRegions.find {
            it.isInRange(row.rowNum, column)
        }
        return row.sheet.getRow(address.firstRow).getCell(address.firstColumn)
    }

    protected PoiCellDefinition createCellIfExists(XSSFCell cell) {
        if (cell) {
            return new PoiCellDefinition(row.sheet.getRowByNumber(cell.rowIndex + 1) ?: row.sheet.createRowWrapper(cell.rowIndex + 1), cell)
        }
        return null
    }

    @Override
    String toString() {
        return "Cell[${row.sheet.name}!${columnAsString}${row.number}]=${value}"
    }

    public <T> T asType(Class<T> type) {
        if (type.isInstance(cell)) {
            return cell as T
        }
        return super.asType(type)
    }
}
