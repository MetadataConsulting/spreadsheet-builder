package builders.dsl.spreadsheet.query.poi;

import builders.dsl.spreadsheet.api.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

class PoiCellStyle implements CellStyle {
    PoiCellStyle(XSSFCellStyle style) {
        this.style = style;
    }

    @Override
    public Color getForeground() {
        if (style.getFillForegroundColorColor() == null) {
            return null;
        }
        if (style.getFillForegroundColorColor().getRGB() == null) {
            return null;
        }
        return new Color(style.getFillForegroundColorColor().getRGB());
    }

    @Override
    public Color getBackground() {
        if (style.getFillBackgroundColorColor() == null) {
            return null;
        }
        if (style.getFillBackgroundColorColor().getRGB() == null) {
            return null;
        }
        return new Color(style.getFillBackgroundColorColor().getRGB());
    }

    @Override
    public ForegroundFill getFill() {
        switch (style.getFillPattern()) {
            case org.apache.poi.ss.usermodel.CellStyle.NO_FILL:
                return ForegroundFill.NO_FILL;
            case org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND:
                return ForegroundFill.SOLID_FOREGROUND;
            case org.apache.poi.ss.usermodel.CellStyle.FINE_DOTS:
                return ForegroundFill.FINE_DOTS;
            case org.apache.poi.ss.usermodel.CellStyle.ALT_BARS:
                return ForegroundFill.ALT_BARS;
            case org.apache.poi.ss.usermodel.CellStyle.SPARSE_DOTS:
                return ForegroundFill.SPARSE_DOTS;
            case org.apache.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS:
                return ForegroundFill.THICK_HORZ_BANDS;
            case org.apache.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS:
                return ForegroundFill.THICK_VERT_BANDS;
            case org.apache.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG:
                return ForegroundFill.THICK_BACKWARD_DIAG;
            case org.apache.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG:
                return ForegroundFill.THICK_FORWARD_DIAG;
            case org.apache.poi.ss.usermodel.CellStyle.BIG_SPOTS:
                return ForegroundFill.BIG_SPOTS;
            case org.apache.poi.ss.usermodel.CellStyle.BRICKS:
                return ForegroundFill.BRICKS;
            case org.apache.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS:
                return ForegroundFill.THIN_HORZ_BANDS;
            case org.apache.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS:
                return ForegroundFill.THIN_VERT_BANDS;
            case org.apache.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG:
                return ForegroundFill.THIN_BACKWARD_DIAG;
            case org.apache.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG:
                return ForegroundFill.THIN_FORWARD_DIAG;
            case org.apache.poi.ss.usermodel.CellStyle.SQUARES:
                return ForegroundFill.SQUARES;
            case org.apache.poi.ss.usermodel.CellStyle.DIAMONDS:
                return ForegroundFill.DIAMONDS;
        }
        return null;
    }

    @Override
    public int getIndent() {
        return style.getIndention();
    }

    @Override
    public int getRotation() {
        return style.getRotation();
    }

    @Override
    public String getFormat() {
        return style.getDataFormatString();
    }

    @Override
    public Font getFont() {
        return style.getFont() != null ? new PoiFont(style.getFont()) : null;
    }

    @Override
    public Border getBorder(Keywords.BorderSide borderSide) {
        if (Keywords.BorderSide.TOP.equals(borderSide)) {
            return new PoiBorder(style.getBorderColor(XSSFCellBorder.BorderSide.TOP), style.getBorderTopEnum());
        } else if (Keywords.BorderSide.BOTTOM.equals(borderSide)) {
            return new PoiBorder(style.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM), style.getBorderBottomEnum());
        } else if (Keywords.BorderSide.LEFT.equals(borderSide)) {
            return new PoiBorder(style.getBorderColor(XSSFCellBorder.BorderSide.LEFT), style.getBorderLeftEnum());
        } else if (Keywords.BorderSide.RIGHT.equals(borderSide)) {
            return new PoiBorder(style.getBorderColor(XSSFCellBorder.BorderSide.RIGHT), style.getBorderRightEnum());
        }
        return new PoiBorder(null, null);
    }

    private final XSSFCellStyle style;
}
