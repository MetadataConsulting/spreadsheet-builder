package builders.dsl.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.FontStyle;
import builders.dsl.spreadsheet.builder.api.FontDefinition;

import java.util.EnumSet;

class PoiFontDefinition implements FontDefinition {

    PoiFontDefinition(XSSFWorkbook workbook) {
        font = workbook.createFont();
    }

    PoiFontDefinition(XSSFWorkbook workbook, XSSFCellStyle style) {
        font = workbook.createFont();
        style.setFont(font);
    }

    @Override
    public PoiFontDefinition color(String hexColor) {
        font.setColor(PoiCellStyleDefinition.parseColor(hexColor));
        return this;
    }

    @Override
    public PoiFontDefinition color(Color colorPreset) {
        color(colorPreset.getHex());
        return this;
    }

    @Override
    public PoiFontDefinition size(int size) {
        font.setFontHeightInPoints((short) size);
        return this;
    }

    @Override
    public PoiFontDefinition name(String name) {
        font.setFontName(name);
        return this;
    }

    @Override
    public PoiFontDefinition style(FontStyle first, FontStyle... other) {
        EnumSet enumSet = EnumSet.of(first, other);
        if (enumSet.contains(FontStyle.ITALIC)){
            font.setItalic(true);
        }


        if (enumSet.contains(FontStyle.BOLD)){
            font.setBold(true);
        }


        if (enumSet.contains(FontStyle.STRIKEOUT)){
            font.setStrikeout(true);
        }


        if (enumSet.contains(FontStyle.UNDERLINE)){
            font.setUnderline(FontUnderline.SINGLE);
        }

        return this;
    }

    XSSFFont getFont() {
        return font;
    }

    private final XSSFFont font;
}
