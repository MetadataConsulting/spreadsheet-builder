package org.modelcatalogue.spreadsheet.query.poi;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.Font;
import org.modelcatalogue.spreadsheet.api.FontStyle;

import java.util.EnumSet;

class PoiFont implements Font {

    PoiFont(XSSFFont xssfFont) {
        this.font = xssfFont;
    }

    @Override
    public Color getColor() {
        if (font.getXSSFColor() == null) {
            return null;
        }
        if (font.getXSSFColor().getRGB() == null) {
            return null;
        }
        return new Color(font.getXSSFColor().getRGB());
    }

    @Override
    public int getSize() {
        return font.getFontHeightInPoints();
    }

    @Override
    public String getName() {
        return font.getFontName();
    }

    @Override
    public EnumSet<FontStyle> getStyles() {
        EnumSet<FontStyle> set = EnumSet.noneOf(FontStyle.class);

        if (font.getItalic()) {
            set.add(FontStyle.ITALIC);
        }


        if (font.getBold()) {
            set.add(FontStyle.BOLD);
        }


        if (font.getStrikeout()) {
            set.add(FontStyle.STRIKEOUT);
        }


        if (font.getUnderline() != FontUnderline.NONE.getByteValue()) {
            set.add(FontStyle.UNDERLINE);
        }

        return set;
    }

    XSSFFont getFont() {
        return font;
    }

    void setFont(XSSFFont font) {
        this.font = font;
    }

    private XSSFFont font;
}
