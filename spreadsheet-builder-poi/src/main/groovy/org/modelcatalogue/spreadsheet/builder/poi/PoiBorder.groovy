package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.modelcatalogue.spreadsheet.api.Border
import org.modelcatalogue.spreadsheet.api.Color

/**
 * Represents the current border configuration of the cell style.
 */
class PoiBorder implements Border {
    private XSSFColor xssfColor
    private BorderStyle borderStyle

    PoiBorder(XSSFColor xssfColor, BorderStyle borderStyle) {
        this.borderStyle = borderStyle
        this.xssfColor = xssfColor
    }

    @Override
    Color getColor() {
        if (xssfColor == null) {
            return null
        }
        return new Color(xssfColor.getRGB())
    }

    @Override
    org.modelcatalogue.spreadsheet.api.BorderStyle getStyle() {
        switch (borderStyle) {
            case BorderStyle.NONE:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.NONE
            case BorderStyle.THIN:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.THIN
            case BorderStyle.MEDIUM:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM
            case BorderStyle.DASHED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASHED
            case BorderStyle.DOTTED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DOTTED
            case BorderStyle.THICK:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.THICK
            case BorderStyle.DOUBLE:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DOUBLE
            case BorderStyle.HAIR:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.HAIR
            case BorderStyle.MEDIUM_DASHED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASHED
            case BorderStyle.DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASH_DOT
            case BorderStyle.MEDIUM_DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASH_DOT
            case BorderStyle.DASH_DOT_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASH_DOT_DOT
            case BorderStyle.MEDIUM_DASH_DOT_DOTC:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASH_DOT_DOT
            case BorderStyle.SLANTED_DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.SLANTED_DASH_DOT
        }
        return null
    }
}
