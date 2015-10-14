package org.modelcatalogue.builder.spreadsheet.api;

public abstract class AbstractCellStyle extends AbstractHTMLColorProvider implements CellStyle {

    @Override
    public ForegroundFill getNoFill() {
        return ForegroundFill.NO_FILL;
    }

    @Override
    public ForegroundFill getSolidForeground() {
        return ForegroundFill.SOLID_FOREGROUND;
    }

    @Override
    public ForegroundFill getFineDots() {
        return ForegroundFill.FINE_DOTS;
    }

    @Override
    public ForegroundFill getAltBars() {
        return ForegroundFill.ALT_BARS;
    }

    @Override
    public ForegroundFill getSparseDots() {
        return ForegroundFill.SPARSE_DOTS;
    }

    @Override
    public ForegroundFill getThickHorizontalBands() {
        return ForegroundFill.THICK_HORZ_BANDS;
    }

    @Override
    public ForegroundFill getThickVerticalBands() {
        return ForegroundFill.THICK_VERT_BANDS;
    }

    @Override
    public ForegroundFill getThickBackwardDiagonals() {
        return ForegroundFill.THICK_BACKWARD_DIAG;
    }

    @Override
    public ForegroundFill getThickForwardDiagonals() {
        return ForegroundFill.THICK_FORWARD_DIAG;
    }

    @Override
    public ForegroundFill getBigSpots() {
        return ForegroundFill.BIG_SPOTS;
    }

    @Override
    public ForegroundFill getBricks() {
        return ForegroundFill.BRICKS;
    }

    @Override
    public ForegroundFill getThinHorizontalBands() {
        return ForegroundFill.THIN_HORZ_BANDS;
    }

    @Override
    public ForegroundFill getThinVerticalBands() {
        return ForegroundFill.THIN_VERT_BANDS;
    }

    @Override
    public ForegroundFill getThinBackwardDiagonals() {
        return ForegroundFill.THIN_BACKWARD_DIAG;
    }

    @Override
    public ForegroundFill getThinForwardDiagonals() {
        return ForegroundFill.THICK_FORWARD_DIAG;
    }

    @Override
    public ForegroundFill getSquares() {
        return ForegroundFill.SQUARES;
    }

    @Override
    public ForegroundFill getDiamonds() {
        return ForegroundFill.DIAMONDS;
    }

    @Override
    public PureBorderSide getLeft() {
        return PureBorderSide.LEFT;
    }

    @Override
    public PureBorderSide getRight() {
        return PureBorderSide.RIGHT;
    }

    @Override
    public BorderSideAndVerticalAlignment getTop() {
        return BorderSideAndVerticalAlignment.TOP;
    }

    @Override
    public BorderSideAndVerticalAlignment getBottom() {
        return BorderSideAndVerticalAlignment.BOTTOM;
    }

    @Override
    public PureVerticalAlignment getCenter() {
        return PureVerticalAlignment.CENTER;
    }

    @Override
    public PureVerticalAlignment getJustify() {
        return PureVerticalAlignment.JUSTIFY;
    }

    @Override
    public PureVerticalAlignment getDistributed() {
        return PureVerticalAlignment.DISTRIBUTED;
    }

    @Override
    public TextKeyword getText() {
        return TextKeyword.WRAP;
    }
}
