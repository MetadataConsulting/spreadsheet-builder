package org.modelcatalogue.spreadsheet.query.poi;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.modelcatalogue.spreadsheet.api.Border;
import org.modelcatalogue.spreadsheet.api.Color;

/**
 * Represents the current border configuration of the cell style.
 */
class PoiBorder implements Border {
    PoiBorder(XSSFColor xssfColor, BorderStyle borderStyle) {

        this.borderStyle = borderStyle;
        this.xssfColor = xssfColor;
    }

    @Override
    public Color getColor() {
        if (xssfColor == null) {
            return null;
        }

        return new Color(xssfColor.getRGB());
    }

    @Override
    public org.modelcatalogue.spreadsheet.api.BorderStyle getStyle() {
        switch (borderStyle) {
            case NONE:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.NONE;
            case THIN:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.THIN;
            case MEDIUM:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM;
            case DASHED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASHED;
            case DOTTED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DOTTED;
            case THICK:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.THICK;
            case DOUBLE:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DOUBLE;
            case HAIR:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.HAIR;
            case MEDIUM_DASHED:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASHED;
            case DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASH_DOT;
            case MEDIUM_DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASH_DOT;
            case DASH_DOT_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.DASH_DOT_DOT;
            case MEDIUM_DASH_DOT_DOTC:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.MEDIUM_DASH_DOT_DOT;
            case SLANTED_DASH_DOT:
                return org.modelcatalogue.spreadsheet.api.BorderStyle.SLANTED_DASH_DOT;
        }
        return null;
    }

    private XSSFColor xssfColor;
    private BorderStyle borderStyle;
}
