package org.modelcatalogue.spreadsheet.builder.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.builder.api.*;
import org.modelcatalogue.spreadsheet.impl.AbstractBorderDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractCellDefinition;
import org.modelcatalogue.spreadsheet.impl.AbstractCellStyleDefinition;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

class PoiCellStyleDefinition extends AbstractCellStyleDefinition {

    PoiCellStyleDefinition(PoiCellDefinition cell) {
        super(cell.getRow().getSheet().getWorkbook());
        if (cell.getCell().getCellStyle().equals(cell.getCell().getSheet().getWorkbook().getStylesSource().getStyleAt(0))) {
            style = cell.getCell().getSheet().getWorkbook().createCellStyle();
            cell.getCell().setCellStyle(style);
        } else {
            style = cell.getCell().getCellStyle();
        }
    }

    PoiCellStyleDefinition(PoiWorkbookDefinition workbook) {
        super(workbook);
        this.style = workbook.getWorkbook().createCellStyle();
    }

    @Override
    protected void doBackground(String hexColor) {
        if (style.getFillForegroundColor() == IndexedColors.AUTOMATIC.getIndex()) {
            foreground(hexColor);
        } else {
            style.setFillBackgroundColor(parseColor(hexColor));
        }
    }

    @Override
    protected void doForeground(String hexColor) {
        if (style.getFillForegroundColor() != IndexedColors.AUTOMATIC.getIndex()) {
            // already set as background color
            style.setFillBackgroundColor(style.getFillForegroundXSSFColor());
        }

        style.setFillForegroundColor(parseColor(hexColor));
        if (style.getFillPatternEnum().equals(FillPatternType.NO_FILL)) {
            fill(ForegroundFill.SOLID_FOREGROUND);
        }
    }

    @Override
    protected void doFill(ForegroundFill fill) {
        switch (fill) {
            case NO_FILL:
                style.setFillPattern(CellStyle.NO_FILL);
                break;
            case SOLID_FOREGROUND:
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                break;
            case FINE_DOTS:
                style.setFillPattern(CellStyle.FINE_DOTS);
                break;
            case ALT_BARS:
                style.setFillPattern(CellStyle.ALT_BARS);
                break;
            case SPARSE_DOTS:
                style.setFillPattern(CellStyle.SPARSE_DOTS);
                break;
            case THICK_HORZ_BANDS:
                style.setFillPattern(CellStyle.THICK_HORZ_BANDS);
                break;
            case THICK_VERT_BANDS:
                style.setFillPattern(CellStyle.THICK_VERT_BANDS);
                break;
            case THICK_BACKWARD_DIAG:
                style.setFillPattern(CellStyle.THICK_BACKWARD_DIAG);
                break;
            case THICK_FORWARD_DIAG:
                style.setFillPattern(CellStyle.THICK_FORWARD_DIAG);
                break;
            case BIG_SPOTS:
                style.setFillPattern(CellStyle.BIG_SPOTS);
                break;
            case BRICKS:
                style.setFillPattern(CellStyle.BRICKS);
                break;
            case THIN_HORZ_BANDS:
                style.setFillPattern(CellStyle.THIN_HORZ_BANDS);
                break;
            case THIN_VERT_BANDS:
                style.setFillPattern(CellStyle.THIN_VERT_BANDS);
                break;
            case THIN_BACKWARD_DIAG:
                style.setFillPattern(CellStyle.THIN_BACKWARD_DIAG);
                break;
            case THIN_FORWARD_DIAG:
                style.setFillPattern(CellStyle.THIN_FORWARD_DIAG);
                break;
            case SQUARES:
                style.setFillPattern(CellStyle.SQUARES);
                break;
            case DIAMONDS:
                style.setFillPattern(CellStyle.DIAMONDS);
                break;
        }
    }

    @Override
    protected FontDefinition createFont() {
        return new PoiFontDefinition(getWorkbook().getWorkbook(), style);
    }

    @Override
    protected void doIndent(int indent) {
        style.setIndention((short) indent);
    }

    @Override
    protected void doWrapText() {
        style.setWrapText(true);
    }

    @Override
    protected void doRotation(int rotation) {
        style.setRotation((short) rotation);
    }

    @Override
    protected void doFormat(String format) {
        XSSFDataFormat dataFormat = getWorkbook().getWorkbook().createDataFormat();
        style.setDataFormat(dataFormat.getFormat(format));
    }

    private PoiWorkbookDefinition getWorkbook() {
        return (PoiWorkbookDefinition) workbook;
    }

    @Override
    protected void doAlign(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment) {
        if (Keywords.VerticalAndHorizontalAlignment.CENTER.equals(verticalAlignment)) {
            style.setVerticalAlignment(VerticalAlignment.CENTER);
        } else if (Keywords.PureVerticalAlignment.DISTRIBUTED.equals(verticalAlignment)) {
            style.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
        } else if (Keywords.VerticalAndHorizontalAlignment.JUSTIFY.equals(verticalAlignment)) {
            style.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        } else if (Keywords.BorderSideAndVerticalAlignment.TOP.equals(verticalAlignment)) {
            style.setVerticalAlignment(VerticalAlignment.TOP);
        } else if (Keywords.BorderSideAndVerticalAlignment.BOTTOM.equals(verticalAlignment)) {
            style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        } else {
            throw new IllegalArgumentException(String.valueOf(verticalAlignment) + " is not supported!");
        }
        if (Keywords.HorizontalAlignment.RIGHT.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else if (Keywords.HorizontalAlignment.LEFT.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if (Keywords.HorizontalAlignment.GENERAL.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.GENERAL);
        } else if (Keywords.HorizontalAlignment.CENTER.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (Keywords.HorizontalAlignment.FILL.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.FILL);
        } else if (Keywords.HorizontalAlignment.JUSTIFY.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.JUSTIFY);
        } else if (Keywords.HorizontalAlignment.CENTER_SELECTION.equals(horizontalAlignment)) {
            style.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        }
    }

    @Override
    protected AbstractBorderDefinition createBorder() {
        return new PoiBorderDefinition(style);
    }

    @Override
    protected void assignTo(AbstractCellDefinition cell) {
        if (cell instanceof PoiCellDefinition) {
            PoiCellDefinition poiCellDefinition = (PoiCellDefinition) cell;
            poiCellDefinition.getCell().setCellStyle(style);
        } else {
            throw new IllegalArgumentException("Cell not supported: " + cell);
        }
    }

    public static XSSFColor parseColor(String hex) {
        if (hex == null) {
            throw new IllegalArgumentException("Please, provide the color in '#abcdef' hex string format");
        }


        Matcher match = Pattern.compile("#([\\dA-F]{2})([\\dA-F]{2})([\\dA-F]{2})").matcher(hex.toUpperCase());

        if (!match.matches()) {
            throw new IllegalArgumentException("Cannot parse color " + hex + ". Please, provide the color in \'#abcdef\' hex string format");
        }


        byte red = (byte) Integer.parseInt(match.group(1), 16);
        byte green = (byte) Integer.parseInt(match.group(2), 16);
        byte blue = (byte) Integer.parseInt(match.group(3), 16);

        return new XSSFColor(new byte[]{red, green, blue});
    }

    void setBorderTo(CellRangeAddress address, PoiSheetDefinition sheet) {
        RegionUtil.setBorderBottom(style.getBorderBottom(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setBorderLeft(style.getBorderLeft(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setBorderRight(style.getBorderRight(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setBorderTop(style.getBorderTop(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setBottomBorderColor(style.getBottomBorderColor(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setLeftBorderColor(style.getLeftBorderColor(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setRightBorderColor(style.getRightBorderColor(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
        RegionUtil.setTopBorderColor(style.getTopBorderColor(), address, sheet.getSheet(), sheet.getSheet().getWorkbook());
    }

    private final XSSFCellStyle style;
}
