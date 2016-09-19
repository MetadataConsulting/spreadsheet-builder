package org.modelcatalogue.builder.spreadsheet.poi

import org.apache.poi.ss.usermodel.FontUnderline
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.modelcatalogue.builder.spreadsheet.api.AbstractHTMLColorProvider
import org.modelcatalogue.builder.spreadsheet.api.Color
import org.modelcatalogue.builder.spreadsheet.api.Font
import org.modelcatalogue.builder.spreadsheet.api.FontStyle

class PoiFont extends AbstractHTMLColorProvider implements Font {

    private final XSSFFont font

    PoiFont(XSSFWorkbook workbook) {
        font = workbook.createFont()
    }

    PoiFont(XSSFWorkbook workbook, XSSFCellStyle style) {
        font = workbook.createFont()
        style.font = font
    }

    @Override
    void color(String hexColor) {
        font.setColor(PoiCellStyle.parseColor(hexColor))
    }

    @Override
    void color(Color colorPreset) {
        color colorPreset.hex
    }

    @Override
    void size(int size) {
        font.setFontHeightInPoints(size.shortValue())
    }

    @Override
    void name(String name) {
        font.setFontName(name)
    }

    @Override
    FontStyle getItalic() {
        FontStyle.ITALIC
    }

    @Override
    FontStyle getBold() {
        FontStyle.BOLD
    }

    @Override
    FontStyle getStrikeout() {
        FontStyle.STRIKEOUT
    }

    @Override
    FontStyle getUnderline() {
        FontStyle.UNDERLINE
    }

    @Override
    void make(FontStyle first, FontStyle... other) {
        EnumSet enumSet = EnumSet.of(first, other)
        if (FontStyle.ITALIC in enumSet) {
            font.italic = true
        }

        if (FontStyle.BOLD in enumSet) {
            font.bold = true
        }

        if (FontStyle.STRIKEOUT in enumSet) {
            font.strikeout = true
        }

        if (FontStyle.UNDERLINE in enumSet) {
            font.setUnderline(FontUnderline.SINGLE)
        }
    }

    protected XSSFFont getFont() {
        return font
    }
}
