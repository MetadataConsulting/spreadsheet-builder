package org.modelcatalogue.spreadsheet.builder.poi

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
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition

import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.api.Configurer
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition
import org.modelcatalogue.spreadsheet.api.ForegroundFill
import org.modelcatalogue.spreadsheet.api.Keywords

import java.util.regex.Matcher

class PoiCellStyleDefinition implements CellStyleDefinition {

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
    PoiCellStyleDefinition base(String stylename) {
        checkSealed()
        Configurer.Runner.doConfigure(workbook.getStyleDefinition(stylename), this)
        this
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
    PoiCellStyleDefinition background(String hexColor) {
        checkSealed()
        if (style.fillForegroundColor == IndexedColors.AUTOMATIC.index) {
            foreground hexColor
        } else {
            style.setFillBackgroundColor(parseColor(hexColor))
        }
        this
    }

    @Override
    PoiCellStyleDefinition background(Color colorPreset) {
        checkSealed()
        background colorPreset.hex
        this
    }

    @Override
    PoiCellStyleDefinition foreground(String hexColor) {
        checkSealed()
        if (style.fillForegroundColor != IndexedColors.AUTOMATIC.index) {
            // already set as background color
            style.setFillBackgroundColor(style.getFillForegroundXSSFColor())
        }
        style.setFillForegroundColor(parseColor(hexColor))
        if (style.fillPatternEnum == FillPatternType.NO_FILL) {
            fill ForegroundFill.SOLID_FOREGROUND
        }
        this
    }

    @Override
    PoiCellStyleDefinition foreground(Color colorPreset) {
        checkSealed()
        foreground colorPreset.hex
        this
    }

    @Override
    PoiCellStyleDefinition fill(ForegroundFill fill) {
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
        this
    }

    @Override
    PoiCellStyleDefinition font(Configurer<FontDefinition> fontConfiguration) {
        checkSealed()
        if (!poiFont) {
            poiFont = new PoiFontDefinition(workbook.workbook, style)
        }
        Configurer.Runner.doConfigure(fontConfiguration, poiFont)
        this
    }

    @Override
    PoiCellStyleDefinition indent(int indent) {
        checkSealed()
        style.indention = (short) indent
        this
    }

    Object getLocked() {
        checkSealed()
        style.locked = true
    }

    @Override
    PoiCellStyleDefinition wrap(Keywords.Text text) {
        checkSealed()
        style.wrapText = true
        this
    }

    Object getHidden() {
        checkSealed()
        style.hidden = true
    }

    @Override
    PoiCellStyleDefinition rotation(int rotation) {
        checkSealed()
        style.rotation = (short) rotation
        this
    }

    @Override
    PoiCellStyleDefinition format(String format) {
        XSSFDataFormat dataFormat = workbook.workbook.createDataFormat()
        style.dataFormat = dataFormat.getFormat(format)
        this
    }

    @Override
    PoiCellStyleDefinition align(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment) {
        checkSealed()
        switch (verticalAlignment) {
            case Keywords.VerticalAndHorizontalAlignment.CENTER:
                style.setVerticalAlignment(VerticalAlignment.CENTER)
                break
            case Keywords.PureVerticalAlignment.DISTRIBUTED:
                style.setVerticalAlignment(VerticalAlignment.DISTRIBUTED)
                break
            case Keywords.VerticalAndHorizontalAlignment.JUSTIFY:
                style.setVerticalAlignment(VerticalAlignment.JUSTIFY)
                break
            case Keywords.BorderSideAndVerticalAlignment.TOP:
                style.setVerticalAlignment(VerticalAlignment.TOP)
                break
            case Keywords.BorderSideAndVerticalAlignment.BOTTOM:
                style.setVerticalAlignment(VerticalAlignment.BOTTOM)
                break
            default:
                throw new IllegalArgumentException("$verticalAlignment is not supported!")
        }
        switch (horizontalAlignment) {
            case Keywords.HorizontalAlignment.RIGHT:
                style.setAlignment(HorizontalAlignment.RIGHT)
                break;
            case Keywords.HorizontalAlignment.LEFT:
                style.setAlignment(HorizontalAlignment.LEFT)
                break;
            case Keywords.HorizontalAlignment.GENERAL:
                style.setAlignment(HorizontalAlignment.GENERAL)
                break;
            case Keywords.HorizontalAlignment.CENTER:
                style.setAlignment(HorizontalAlignment.CENTER)
                break;
            case Keywords.HorizontalAlignment.FILL:
                style.setAlignment(HorizontalAlignment.FILL)
                break;
            case Keywords.HorizontalAlignment.JUSTIFY:
                style.setAlignment(HorizontalAlignment.JUSTIFY)
                break;
            case Keywords.HorizontalAlignment.CENTER_SELECTION:
                style.setAlignment(HorizontalAlignment.CENTER_SELECTION)
                break;
        }
        return this
    }

    @Override
    PoiCellStyleDefinition border(Configurer<BorderDefinition> borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder)

        Keywords.BorderSide.BORDER_SIDES.each { Keywords.BorderSide side ->
            poiBorder.applyTo(side)
        }
        this
    }

    @Override
    PoiCellStyleDefinition border(Keywords.BorderSide location, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder)
        poiBorder.applyTo(location)
        this
    }

    @Override
    PoiCellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second,
                Configurer<BorderDefinition> borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder)
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)
        this
    }

    @Override
    PoiCellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third,
                Configurer<BorderDefinition> borderConfiguration) {
        checkSealed()
        PoiBorderDefinition poiBorder = findOrCreateBorder()
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder)
        poiBorder.applyTo(first)
        poiBorder.applyTo(second)
        poiBorder.applyTo(third)
        this
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
