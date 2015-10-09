package org.modelcatalogue.builder.spreadsheet.poi

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.modelcatalogue.builder.spreadsheet.api.AbstractBorder
import org.modelcatalogue.builder.spreadsheet.api.BorderSide
import org.modelcatalogue.builder.spreadsheet.api.BorderSideAndVerticalAlignment
import org.modelcatalogue.builder.spreadsheet.api.BorderStyle
import org.modelcatalogue.builder.spreadsheet.api.Color
import org.modelcatalogue.builder.spreadsheet.api.PureBorderSide

@CompileStatic class PoiBorder extends AbstractBorder {

    private final XSSFCellStyle xssfCellStyle

    private XSSFColor color
    private BorderStyle borderStyle

    PoiBorder(XSSFCellStyle xssfCellStyle) {
        this.xssfCellStyle = xssfCellStyle
    }

    @Override
    void style(BorderStyle style) {
        borderStyle = style
    }

    @Override
    void color(String hexColor) {
        color = PoiCellStyle.parseColor(hexColor)
    }

    @Override
    void color(Color colorPreset) {
        color colorPreset.hex
    }

    protected void applyTo(BorderSide location) {
        switch (location) {
            case PureBorderSide.BOTTOM:
                if (borderStyle) {
                    xssfCellStyle.setBorderBottom(poiBorderStyle)
                }
                if (color) {
                    xssfCellStyle.setBottomBorderColor(color)
                }
                break
            case PureBorderSide.TOP:
                if (borderStyle) {
                    xssfCellStyle.setBorderTop(poiBorderStyle)
                }
                if (color) {
                    xssfCellStyle.setTopBorderColor(color)
                }
                break
            case BorderSideAndVerticalAlignment.LEFT:
                if (borderStyle) {
                    xssfCellStyle.setBorderLeft(poiBorderStyle)
                }
                if (color) {
                    xssfCellStyle.setLeftBorderColor(color)
                }
                break
            case BorderSideAndVerticalAlignment.RIGHT:
                if (borderStyle) {
                    xssfCellStyle.setBorderRight(poiBorderStyle)
                }
                if (color) {
                    xssfCellStyle.setRightBorderColor(color)
                }
                break
            default:
                throw new IllegalArgumentException("$location is not supported!")
        }

    }

    private org.apache.poi.ss.usermodel.BorderStyle getPoiBorderStyle() {
        switch (borderStyle) {
            case BorderStyle.NONE:
                return org.apache.poi.ss.usermodel.BorderStyle.NONE
            case BorderStyle.THIN:
                return org.apache.poi.ss.usermodel.BorderStyle.THIN
            case BorderStyle.MEDIUM:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM
            case BorderStyle.DASHED:
                return org.apache.poi.ss.usermodel.BorderStyle.DASHED
            case BorderStyle.DOTTED:
                return org.apache.poi.ss.usermodel.BorderStyle.DOTTED
            case BorderStyle.THICK:
                return org.apache.poi.ss.usermodel.BorderStyle.THICK
            case BorderStyle.DOUBLE:
                return org.apache.poi.ss.usermodel.BorderStyle.DOUBLE
            case BorderStyle.HAIR:
                return org.apache.poi.ss.usermodel.BorderStyle.HAIR
            case BorderStyle.MEDIUM_DASHED:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASHED
            case BorderStyle.DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT
            case BorderStyle.MEDIUM_DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT
            case BorderStyle.DASH_DOT_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT_DOT
            case BorderStyle.MEDIUM_DASH_DOT_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT_DOTC
            case BorderStyle.SLANTED_DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.SLANTED_DASH_DOT
        }
    }

}
