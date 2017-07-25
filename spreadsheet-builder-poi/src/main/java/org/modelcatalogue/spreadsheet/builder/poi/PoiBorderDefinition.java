package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractBorderDefinition;

class PoiBorderDefinition extends AbstractBorderDefinition {

    PoiBorderDefinition(XSSFCellStyle xssfCellStyle) {
        this.xssfCellStyle = xssfCellStyle;
    }

    @Override
    public BorderDefinition color(String hexColor) {
        color = PoiCellStyleDefinition.parseColor(hexColor);
        return this;
    }

    protected void applyTo(Keywords.BorderSide location) {
        org.apache.poi.ss.usermodel.BorderStyle poiBorderStyle = getPoiBorderStyle();
        if (Keywords.BorderSideAndHorizontalAlignment.BOTTOM.equals(location)) {
            if (poiBorderStyle != null) {
                xssfCellStyle.setBorderBottom(poiBorderStyle);
            }

            if (color != null) {
                xssfCellStyle.setBottomBorderColor(color);
            }

        } else if (Keywords.BorderSideAndHorizontalAlignment.TOP.equals(location)) {
            if (poiBorderStyle != null) {
                xssfCellStyle.setBorderTop(poiBorderStyle);
            }

            if (color != null) {
                xssfCellStyle.setTopBorderColor(color);
            }

        } else if (Keywords.BorderSideAndVerticalAlignment.LEFT.equals(location)) {
            if (poiBorderStyle != null) {
                xssfCellStyle.setBorderLeft(poiBorderStyle);
            }

            if (color != null) {
                xssfCellStyle.setLeftBorderColor(color);
            }

        } else if (Keywords.BorderSideAndVerticalAlignment.RIGHT.equals(location)) {
            if (poiBorderStyle != null) {
                xssfCellStyle.setBorderRight(poiBorderStyle);
            }

            if (color != null) {
                xssfCellStyle.setRightBorderColor(color);
            }

        } else {
            throw new IllegalArgumentException(String.valueOf(location) + " is not supported!");
        }

    }

    private org.apache.poi.ss.usermodel.BorderStyle getPoiBorderStyle() {
        if (borderStyle == null) {
            return null;
        }
        switch (borderStyle) {
            case NONE:
                return org.apache.poi.ss.usermodel.BorderStyle.NONE;
            case THIN:
                return org.apache.poi.ss.usermodel.BorderStyle.THIN;
            case MEDIUM:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM;
            case DASHED:
                return org.apache.poi.ss.usermodel.BorderStyle.DASHED;
            case DOTTED:
                return org.apache.poi.ss.usermodel.BorderStyle.DOTTED;
            case THICK:
                return org.apache.poi.ss.usermodel.BorderStyle.THICK;
            case DOUBLE:
                return org.apache.poi.ss.usermodel.BorderStyle.DOUBLE;
            case HAIR:
                return org.apache.poi.ss.usermodel.BorderStyle.HAIR;
            case MEDIUM_DASHED:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASHED;
            case DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT;
            case MEDIUM_DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT;
            case DASH_DOT_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT_DOT;
            case MEDIUM_DASH_DOT_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT_DOT;
            case SLANTED_DASH_DOT:
                return org.apache.poi.ss.usermodel.BorderStyle.SLANTED_DASH_DOT;
        }
        throw new IllegalStateException("Uknown style: " + borderStyle);
    }

    private final XSSFCellStyle xssfCellStyle;
    private XSSFColor color;
}
