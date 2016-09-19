package org.modelcatalogue.spreadsheet.builder.poi

import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.modelcatalogue.spreadsheet.api.CellStyle
import org.modelcatalogue.spreadsheet.api.Color
import org.modelcatalogue.spreadsheet.api.ForegroundFill

class PoiCellStyle implements CellStyle {

    private final XSSFCellStyle style

    PoiCellStyle(XSSFCellStyle style) {
        this.style = style
    }

    @Override
    Color getForeground() {
        return style.getFillForegroundColorColor()?.getRGB() ? new Color(style.getFillForegroundColorColor().getRGB()) : null
    }

    @Override
    Color getBackground() {
        return style.getFillBackgroundColorColor()?.getRGB() ? new Color(style.getFillBackgroundColorColor().getRGB()) : null
    }

    @Override
    ForegroundFill getFill() {
        switch (style.fillPattern) {
            case org.apache.poi.ss.usermodel.CellStyle.NO_FILL: return ForegroundFill.NO_FILL
            case org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND: return ForegroundFill.SOLID_FOREGROUND
            case org.apache.poi.ss.usermodel.CellStyle.FINE_DOTS: return ForegroundFill.FINE_DOTS
            case org.apache.poi.ss.usermodel.CellStyle.ALT_BARS: return ForegroundFill.ALT_BARS
            case org.apache.poi.ss.usermodel.CellStyle.SPARSE_DOTS: return ForegroundFill.SPARSE_DOTS
            case org.apache.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS: return ForegroundFill.THICK_HORZ_BANDS
            case org.apache.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS: return ForegroundFill.THICK_VERT_BANDS
            case org.apache.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG: return ForegroundFill.THICK_BACKWARD_DIAG
            case org.apache.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG: return ForegroundFill.THICK_FORWARD_DIAG
            case org.apache.poi.ss.usermodel.CellStyle.BIG_SPOTS: return ForegroundFill.BIG_SPOTS
            case org.apache.poi.ss.usermodel.CellStyle.BRICKS: return ForegroundFill.BRICKS
            case org.apache.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS: return ForegroundFill.THIN_HORZ_BANDS
            case org.apache.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS: return ForegroundFill.THIN_VERT_BANDS
            case org.apache.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG: return ForegroundFill.THIN_BACKWARD_DIAG
            case org.apache.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG: return ForegroundFill.THIN_FORWARD_DIAG
            case org.apache.poi.ss.usermodel.CellStyle.SQUARES: return ForegroundFill.SQUARES
            case org.apache.poi.ss.usermodel.CellStyle.DIAMONDS: return ForegroundFill.DIAMONDS
        }
        return null
    }

    @Override
    int getIndent() {
        return style.getIndention()
    }

    @Override
    int getRotation() {
        return style.getRotation()
    }

    @Override
    String getFormat() {
        return style.dataFormatString
    }
}
