package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.ss.usermodel.FontUnderline
import org.apache.poi.xssf.usermodel.XSSFFont
import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.api.Font
import org.modelcatalogue.spreadsheet.api.FontStyle

class PoiFont implements Font {
    XSSFFont font

    PoiFont(XSSFFont xssfFont) {
        this.font = xssfFont
    }

    @Override
    Color getColor() {
        return font.getXSSFColor()?.getRGB() ? new Color(font.getXSSFColor().getRGB()) : null
    }

    @Override
    int getSize() {
        return font.getFontHeightInPoints()
    }

    @Override
    String getName() {
        return font.getFontName()
    }

    @Override
    EnumSet<FontStyle> getStyles() {
        EnumSet<FontStyle> set = EnumSet.noneOf(FontStyle)

        if (font.italic) {
            set.add(FontStyle.ITALIC)
        }

        if (font.bold) {
            set.add(FontStyle.BOLD)
        }

        if (font.strikeout) {
            set.add(FontStyle.STRIKEOUT)
        }

        if (font.underline != FontUnderline.NONE.byteValue) {
            set.add(FontStyle.UNDERLINE)
        }

        return set
    }
}
