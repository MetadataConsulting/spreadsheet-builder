package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.*;

public abstract class AbstractCellStyleDefinition implements CellStyleDefinition, HTMLColors {

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
    public Keywords.PureBorderSide getLeft() {
        return Keywords.PureBorderSide.LEFT;
    }

    @Override
    public Keywords.PureBorderSide getRight() {
        return Keywords.PureBorderSide.RIGHT;
    }

    @Override
    public Keywords.BorderSideAndVerticalAlignment getTop() {
        return Keywords.BorderSideAndVerticalAlignment.TOP;
    }

    @Override
    public Keywords.BorderSideAndVerticalAlignment getBottom() {
        return Keywords.BorderSideAndVerticalAlignment.BOTTOM;
    }

    @Override
    public Keywords.PureVerticalAlignment getCenter() {
        return Keywords.PureVerticalAlignment.CENTER;
    }

    @Override
    public Keywords.PureVerticalAlignment getJustify() {
        return Keywords.PureVerticalAlignment.JUSTIFY;
    }

    @Override
    public Keywords.PureVerticalAlignment getDistributed() {
        return Keywords.PureVerticalAlignment.DISTRIBUTED;
    }

    @Override
    public Keywords.Text getText() {
        return Keywords.Text.WRAP;
    }
}
