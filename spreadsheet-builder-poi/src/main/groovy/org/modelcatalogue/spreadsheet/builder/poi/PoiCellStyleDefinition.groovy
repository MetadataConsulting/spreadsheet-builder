package org.modelcatalogue.spreadsheet.builder.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.modelcatalogue.spreadsheet.api.HTMLColors
import org.modelcatalogue.spreadsheet.builder.api.AbstractCellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition

import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition
import org.modelcatalogue.spreadsheet.api.ForegroundFill
import org.modelcatalogue.spreadsheet.builder.api.Keywords

import org.modelcatalogue.spreadsheet.builder.api.HorizontalAlignmentConfigurer

import java.util.regex.Matcher

class PoiCellStyleDefinition extends AbstractCellStyleDefinition implements HTMLColors {

    private final XSSFCellStyle style
    private final PoiWorkbookDefinition workbook

    private PoiFontDefinition poiFont
    protected PoiBorderDefinition poiBorder

    private boolean sealed

    PoiCellStyleDefinition(PoiCellDefinition cell) {
        if (cell.cell.cellStyle == cell.cell.sheet.workbook.stylesSource.getStyleAt(0)) {
            style = cell.cell.sheet.workbook.createCellStyle() as XSSFCellStyle
            cell.cell.cellStyle = style
        } else {
            style = cell.cell.cellStyle as XSSFCellStyle
        }
        workbook = cell.row.sheet.workbook
    }

    PoiCellStyleDefinition(PoiWorkbookDefinition workbook) {
        this.style = workbook.workbook.createCellStyle() as XSSFCellStyle
        this.workbook = workbook
    }

    @Override
    void base(String stylename) {
        checkSealed()
        with workbook.getStyleDefinition(stylename)
    }

    void checkSealed() {
        if (sealed) {
            throw new IllegalStateException("The cell style is already sealed! You need to create new style. Use 'styles' method to combine multiple named styles! Create new named style if you're trying to update existing style with closure definition.")
        }
    }

    boolean isSealed() {
        return sealed
    }

    @Override
    void background(String hexColor) {
        checkSealed()
        if (style.fillForegroundColor == IndexedColors.AUTOMATIC.index) {
            foreground hexColor
        } else {
            style.setFillBackgroundColor(parseColor(hexColor))
        }
    }

    @Override
    void background(Color colorPreset) {
        checkSealed()
        background colorPreset.hex
    }

    @Override
    void foreground(String hexColor) {
        checkSealed()
        if (style.fillForegroundColor != IndexedColors.AUTOMATIC.index) {
            // already set as background color
            style.setFillBackgroundColor(style.getFillForegroundXSSFColor())
        }
        style.setFillForegroundColor(parseColor(hexColor))
        if (style.fillPatternEnum == FillPatternType.NO_FILL) {
            fill solidForeground
        }
    }

    @Override
    void foreground(Color colorPreset) {
        checkSealed()
        foreground colorPreset.hex
    }

    @Override
    void fill(ForegroundFill fill) {
        checkSealed()
        switch (fill) {
            case ForegroundFill.NO_FILL:
                style.fillPattern = CellStyle.NO_FILL
                break
            case ForegroundFill.SOLID_FOREGROUND:
                style.fillPattern = CellStyle.SOLID_FOREGROUND
                break
            case ForegroundFill.FINE_DOTS:
                style.fillPattern = CellStyle.FINE_DOTS
                break
            case ForegroundFill.ALT_BARS:
                style.fillPattern = CellStyle.ALT_BARS
                break
            case ForegroundFill.SPARSE_DOTS:
                style.fillPattern = CellStyle.SPARSE_DOTS
                break
            case ForegroundFill.THICK_HORZ_BANDS:
                style.fillPattern = CellStyle.THICK_HORZ_BANDS
                break
            case ForegroundFill.THICK_VERT_BANDS:
                style.fillPattern = CellStyle.THICK_VERT_BANDS
                break
            case ForegroundFill.THICK_BACKWARD_DIAG:
                style.fillPattern = CellStyle.THICK_BACKWARD_DIAG
                break
            case ForegroundFill.THICK_FORWARD_DIAG:
                style.fillPattern = CellStyle.THICK_FORWARD_DIAG
                break
            case ForegroundFill.BIG_SPOTS:
                style.fillPattern = CellStyle.BIG_SPOTS
                break
            case ForegroundFill.BRICKS:
                style.fillPattern = CellStyle.BRICKS
                break
            case ForegroundFill.THIN_HORZ_BANDS:
                style.fillPattern = CellStyle.THIN_HORZ_BANDS
                break
            case ForegroundFill.THIN_VERT_BANDS:
                style.fillPattern = CellStyle.THIN_VERT_BANDS
                break
            case ForegroundFill.THIN_BACKWARD_DIAG:
                style.fillPattern = CellStyle.THIN_BACKWARD_DIAG
                break
            case ForegroundFill.THIN_FORWARD_DIAG:
                style.fillPattern = CellStyle.THIN_FORWARD_DIAG
                break
            case ForegroundFill.SQUARES:
                style.fillPattern = CellStyle.SQUARES
                break
            case ForegroundFill.DIAMONDS:
                style.fillPattern = CellStyle.DIAMONDS
                break
        }
    }

    @Override
    void font(@DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.FontDefinition") Closure fontConfiguration) {
        checkSealed()
        if (!poiFont) {
            poiFont = new PoiFontDefinition(workbook.workbook, style)
        }
        poiFont.with fontConfiguration
    }

    @Override
    void indent(int indent) {
        checkSealed()
        style.indention = (short) indent
    }

    Object getLocked() {
        checkSealed()
        style.locked = true
    }

    @Override
    void wrap(Keywords.Text text) {
        checkSealed()
        style.wrapText = true
    }

    Object getHidden() {
        checkSealed()
        style.hidden = true
    }

    @Override
    void rotation(int rotation) {
        checkSealed()
        style.rotation = (short) rotation
    }

    @Override
    void format(String format) {
        XSSFDataFormat dataFormat = workbook.workbook.createDataFormat()
        style.dataFormat = dataFormat.getFormat(format)
    }

    @Override
    HorizontalAlignmentConfigurer align(Keywords.VerticalAlignment alignment) {
        checkSealed()
        switch (alignment) {
            case Keywords.PureVerticalAlignment.CENTER:
                style.setVerticalAlignment(VerticalAlignment.CENTER)
                break
            case Keywords.PureVerticalAlignment.DISTRIBUTED:
                style.setVerticalAlignment(VerticalAlignment.DISTRIBUTED)
                break
            case Keywords.PureVerticalAlignment.JUSTIFY:
                style.setVerticalAlignment(VerticalAlignment.JUSTIFY)
                break
            case Keywords.BorderSideAndVerticalAlignment.TOP:
                style.setVerticalAlignment(VerticalAlignment.TOP)
                break
            case Keywords.BorderSideAndVerticalAlignment.BOTTOM:
                style.setVerticalAlignment(VerticalAlignment.BOTTOM)
                break
            default:
                throw new IllegalArgumentException("$alignment is not supported!")
        }
        return new PoiHorizontalAlignmentConfigurer(this)
    }

    @Override
    void border(@DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration

        Keywords.BorderSide.BORDER_SIDES.each { Keywords.BorderSide side ->
            poiBorder.applyTo(side)
        }
    }

    @Override
    void border(Keywords.BorderSide location, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(location)
    }

    @Override
    void border(Keywords.BorderSide first, Keywords.BorderSide second,
                @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)

    }

    @Override
    void border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third,
                @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)
        poiBorder.applyTo(third)
    }

    private PoiBorderDefinition findOrCreateBorder() {
        if (!poiBorder) {
            poiBorder = new PoiBorderDefinition(style)
        }
        poiBorder
    }

    protected setHorizontalAlignment(HorizontalAlignment alignment) {
        style.setAlignment(alignment)
    }

    static XSSFColor parseColor(String hex) {
        if (!hex) {
            throw new IllegalArgumentException("Please, provide the color in '#abcdef' hex string format")
        }
        Matcher match = hex.toUpperCase() =~ /#([\dA-F]{2})([\dA-F]{2})([\dA-F]{2})/

        if (!match.matches()) {
            throw new IllegalArgumentException("Cannot parse color $hex. Please, provide the color in '#abcdef' hex string format")
        }


        byte red = Integer.parseInt(match.group(1), 16) as byte
        byte green = Integer.parseInt(match.group(2), 16) as byte
        byte blue = Integer.parseInt(match.group(3), 16) as byte

        new XSSFColor([red, green, blue] as byte[])
    }

    void seal() {
        this.sealed = true
    }


    void assignTo(PoiCellDefinition cell) {
        cell.cell.cellStyle = style
    }

    void setBorderTo(CellRangeAddress address, PoiSheetDefinition sheet) {
        RegionUtil.setBorderBottom(style.borderBottom, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setBorderLeft(style.borderLeft, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setBorderRight(style.borderRight, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setBorderTop(style.borderTop, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setBottomBorderColor(style.bottomBorderColor, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setLeftBorderColor(style.leftBorderColor, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setRightBorderColor(style.rightBorderColor, address, sheet.sheet, sheet.sheet.workbook)
        RegionUtil.setTopBorderColor(style.topBorderColor, address, sheet.sheet, sheet.sheet.workbook)
    }
}
