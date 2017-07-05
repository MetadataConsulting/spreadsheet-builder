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
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition;
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition;
import org.modelcatalogue.spreadsheet.builder.api.FontDefinition;
import org.modelcatalogue.spreadsheet.builder.api.Sealable;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

class PoiCellStyleDefinition implements CellStyleDefinition, Sealable {

    PoiCellStyleDefinition(PoiCellDefinition cell) {
        if (cell.getCell().getCellStyle().equals(cell.getCell().getSheet().getWorkbook().getStylesSource().getStyleAt(0))) {
            style = cell.getCell().getSheet().getWorkbook().createCellStyle();
            cell.getCell().setCellStyle(style);
        } else {
            style = cell.getCell().getCellStyle();
        }

        workbook = cell.getRow().getSheet().getWorkbook();
    }

    PoiCellStyleDefinition(PoiWorkbookDefinition workbook) {
        this.style = workbook.getWorkbook().createCellStyle();
        this.workbook = workbook;
    }

    @Override
    public PoiCellStyleDefinition base(String stylename) {
        checkSealed();
        Configurer.Runner.doConfigure(workbook.getStyleDefinition(stylename), this);
        return this;
    }

    public void checkSealed() {
        if (sealed) {
            throw new IllegalStateException("The cell style is already sealed! You need to create new style. Use 'styles' method to combine multiple named styles! Create new named style if you're trying to update existing style with closure definition.");
        }

    }

    public boolean isSealed() {
        return sealed;
    }

    @Override
    public PoiCellStyleDefinition background(String hexColor) {
        checkSealed();
        if (style.getFillForegroundColor() == IndexedColors.AUTOMATIC.getIndex()) {
            foreground(hexColor);
        } else {
            style.setFillBackgroundColor(parseColor(hexColor));
        }

        return this;
    }

    @Override
    public PoiCellStyleDefinition background(Color colorPreset) {
        checkSealed();
        background(colorPreset.getHex());
        return this;
    }

    @Override
    public PoiCellStyleDefinition foreground(String hexColor) {
        checkSealed();
        if (style.getFillForegroundColor() != IndexedColors.AUTOMATIC.getIndex()) {
            // already set as background color
            style.setFillBackgroundColor(style.getFillForegroundXSSFColor());
        }

        style.setFillForegroundColor(parseColor(hexColor));
        if (style.getFillPatternEnum().equals(FillPatternType.NO_FILL)) {
            fill(ForegroundFill.SOLID_FOREGROUND);
        }

        return this;
    }

    @Override
    public PoiCellStyleDefinition foreground(Color colorPreset) {
        checkSealed();
        foreground(colorPreset.getHex());
        return this;
    }

    @Override
    public PoiCellStyleDefinition fill(ForegroundFill fill) {
        checkSealed();
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
        return this;
    }

    @Override
    public PoiCellStyleDefinition font(Configurer<FontDefinition> fontConfiguration) {
        checkSealed();
        if (poiFont == null) {
            poiFont = new PoiFontDefinition(workbook.getWorkbook(), style);
        }

        Configurer.Runner.doConfigure(fontConfiguration, poiFont);
        return this;
    }

    @Override
    public PoiCellStyleDefinition indent(int indent) {
        checkSealed();
        style.setIndention((short) indent);
        return this;
    }

    public Object getLocked() {
        checkSealed();
        return setLocked(style, true);
    }

    @Override
    public PoiCellStyleDefinition wrap(Keywords.Text text) {
        checkSealed();
        style.setWrapText(true);
        return this;
    }

    public Object getHidden() {
        checkSealed();
        return setHidden(style, true);
    }

    @Override
    public PoiCellStyleDefinition rotation(int rotation) {
        checkSealed();
        style.setRotation((short) rotation);
        return this;
    }

    @Override
    public PoiCellStyleDefinition format(String format) {
        XSSFDataFormat dataFormat = workbook.getWorkbook().createDataFormat();
        style.setDataFormat(dataFormat.getFormat(format));
        return this;
    }

    @Override
    public PoiCellStyleDefinition align(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment) {
        checkSealed();
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
        return this;
    }

    @Override
    public PoiCellStyleDefinition border(Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        final PoiBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);

        for (Keywords.BorderSide side: Keywords.BorderSide.BORDER_SIDES) {
            poiBorder.applyTo(side);
        }
        return this;
    }

    @Override
    public PoiCellStyleDefinition border(Keywords.BorderSide location, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        PoiBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(location);
        return this;
    }

    @Override
    public PoiCellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        PoiBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(first);
        poiBorder.applyTo(second);
        return this;
    }

    @Override
    public PoiCellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        PoiBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(first);
        poiBorder.applyTo(second);
        poiBorder.applyTo(third);
        return this;
    }

    private PoiBorderDefinition findOrCreateBorder() {
        if (poiBorder == null) {
            poiBorder = new PoiBorderDefinition(style);
        }

        return poiBorder;
    }

    protected void setHorizontalAlignment(HorizontalAlignment alignment) {
        style.setAlignment(alignment);
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

    public void seal() {
        this.sealed = true;
    }

    public void assignTo(PoiCellDefinition cell) {
        cell.getCell().setCellStyle(style);
    }

    public void setBorderTo(CellRangeAddress address, PoiSheetDefinition sheet) {
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
    private final PoiWorkbookDefinition workbook;
    private PoiFontDefinition poiFont;
    private PoiBorderDefinition poiBorder;
    private boolean sealed;

    private static boolean setLocked(XSSFCellStyle propOwner, boolean locked) {
        propOwner.setLocked(locked);
        return locked;
    }

    private static boolean setHidden(XSSFCellStyle propOwner, boolean hidden) {
        propOwner.setHidden(hidden);
        return hidden;
    }
}
