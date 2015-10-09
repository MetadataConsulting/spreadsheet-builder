package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFDataFormat
import org.apache.poi.xssf.usermodel.XSSFWorkbook
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

@CompileStatic class PoiCellStyle extends AbstractCellStyle {

    private final XSSFCellStyle style
    private final XSSFWorkbook workbook

    private PoiFont poiFont
    private PoiBorder poiBorder

    PoiCellStyle(XSSFCell xssfCell) {
        workbook = xssfCell.row.sheet.workbook as XSSFWorkbook
        if (xssfCell.cellStyle == workbook.stylesSource.getStyleAt(0)) {
            style = workbook.createCellStyle() as XSSFCellStyle
            xssfCell.cellStyle = style
        } else {
            style = xssfCell.cellStyle as XSSFCellStyle
        }
    }

    @Override
    void background(String hexColor) {
        if (style.fillForegroundColor == IndexedColors.AUTOMATIC.index) {
            foreground hexColor
        } else {
            style.setFillBackgroundColor(parseColor(hexColor))
        }
    }

    @Override
    void background(Color colorPreset) {
        background colorPreset.hex
    }

    @Override
    void foreground(String hexColor) {
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
        foreground colorPreset.hex
    }

    @Override
    void fill(ForegroundFill fill) {
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
    void font(@DelegatesTo(Font.class) Closure fontConfiguration) {
        if (!poiFont) {
            poiFont = new PoiFont(workbook, style)
        }
        poiFont.with fontConfiguration
    }

    @Override
    void indent(int indent) {
        style.indention = (short) indent
    }

    Object getLocked() {
        style.locked = true
    }

    @Override
    void wrap(TextKeyword text) {
        style.wrapText = true
    }

    Object getHidden() {
        style.hidden = true
    }

    @Override
    void rotation(int rotation) {
        style.rotation = (short) rotation
    }

    @Override
    void format(String format) {
        XSSFDataFormat dataFormat = workbook.createDataFormat()
        style.dataFormat = dataFormat.getFormat(format)
    }

    @Override
    HorizontalAlignmentConfigurer align(VerticalAlignment alignment) {
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
    void border(@DelegatesTo(Border.class) Closure borderConfiguration) {
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration

        BorderSide.BORDER_SIDES.each { BorderSide side ->
            poiBorder.applyTo(side)
        }
    }

    @Override
    void border(BorderSide location, @DelegatesTo(Border.class) Closure borderConfiguration) {
        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(location)
    }

    @Override
    void border(BorderSide first, BorderSide second,
                @DelegatesTo(Border.class) Closure borderConfiguration) {

        PoiBorder poiBorder = findOrCreateBorder()
        poiBorder.with borderConfiguration
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)

    }

    @Override
    void border(BorderSide first, BorderSide second, BorderSide third,
                @DelegatesTo(Border.class) Closure borderConfiguration) {

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


    protected setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment alignment) {
        style.setVerticalAlignment(alignment)
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
}
