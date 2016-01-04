package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.stc.ClosureParams
import groovy.transform.stc.FromString
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.modelcatalogue.builder.spreadsheet.api.AbstractCellStyle
import org.modelcatalogue.builder.spreadsheet.api.Border
import org.modelcatalogue.builder.spreadsheet.api.BorderSide
import org.modelcatalogue.builder.spreadsheet.api.BorderSideAndVerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.Color
import org.modelcatalogue.builder.spreadsheet.api.Font
import org.modelcatalogue.builder.spreadsheet.api.ForegroundFill
import org.modelcatalogue.builder.spreadsheet.api.VerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.PureVerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.HorizontalAlignmentConfigurer
import org.modelcatalogue.builder.spreadsheet.api.TextKeyword

import java.util.regex.Matcher

class PoiCellStyle extends AbstractCellStyle {

    private final XSSFCellStyle style
    private final PoiWorkbook workbook

    private PoiFont poiFont
    private PoiBorder poiBorder

    private boolean sealed

    PoiCellStyle(PoiCell cell) {
        if (cell.cell.cellStyle == cell.cell.sheet.workbook.stylesSource.getStyleAt(0)) {
            style = cell.cell.sheet.workbook.createCellStyle() as XSSFCellStyle
            cell.cell.cellStyle = style
        } else {
            style = cell.cell.cellStyle as XSSFCellStyle
        }
        workbook = cell.row.sheet.workbook
    }

    PoiCellStyle(PoiWorkbook workbook) {
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
            throw new IllegalStateException("The cell style is already sealed! You need to create new style.")
        }
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
    void font(@DelegatesTo(Font.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Font") Closure fontConfiguration) {
        checkSealed()
        if (!poiFont) {
            poiFont = new PoiFont(workbook.workbook, style)
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
    void wrap(TextKeyword text) {
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
    HorizontalAlignmentConfigurer align(VerticalAlignment alignment) {
        checkSealed()
        switch (alignment) {
            case PureVerticalAlignment.CENTER:
                style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER)
                break
            case PureVerticalAlignment.DISTRIBUTED:
                style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.DISTRIBUTED)
                break
            case PureVerticalAlignment.JUSTIFY:
                style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.JUSTIFY)
                break
            case BorderSideAndVerticalAlignment.TOP:
                style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP)
                break
            case BorderSideAndVerticalAlignment.BOTTOM:
                style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM)
                break
            default:
                throw new IllegalArgumentException("$alignment is not supported!")
        }
        return new PoiHorizontalAlignmentConfigurer(this)
    }

    @Override
    void border(@DelegatesTo(Border.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Border") Closure borderConfiguration) {
        checkSealed()
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration

        BorderSide.BORDER_SIDES.each { BorderSide side ->
            poiBorder.applyTo(side)
        }
    }

    @Override
    void border(BorderSide location, @DelegatesTo(Border.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Border") Closure borderConfiguration) {
        checkSealed()
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(location)
    }

    @Override
    void border(BorderSide first, BorderSide second,
                @DelegatesTo(Border.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Border") Closure borderConfiguration) {
        checkSealed()
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)

    }

    @Override
    void border(BorderSide first, BorderSide second, BorderSide third,
                @DelegatesTo(Border.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Border") Closure borderConfiguration) {
        checkSealed()
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)
        poiBorder.applyTo(third)
    }

    private PoiBorder findOrCreateBorder() {
        if (!poiBorder) {
            poiBorder = new PoiBorder(style)
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
}
