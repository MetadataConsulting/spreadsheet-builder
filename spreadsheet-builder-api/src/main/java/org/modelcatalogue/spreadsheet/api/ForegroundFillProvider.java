package org.modelcatalogue.spreadsheet.api;

public interface ForegroundFillProvider {
    ForegroundFill getNoFill();

    ForegroundFill getSolidForeground();

    ForegroundFill getFineDots();

    ForegroundFill getAltBars();

    ForegroundFill getSparseDots();

    ForegroundFill getThickHorizontalBands();

    ForegroundFill getThickVerticalBands();

    ForegroundFill getThickBackwardDiagonals();

    ForegroundFill getThickForwardDiagonals();

    ForegroundFill getBigSpots();

    ForegroundFill getBricks();

    ForegroundFill getThinHorizontalBands();

    ForegroundFill getThinVerticalBands();

    ForegroundFill getThinBackwardDiagonals();

    ForegroundFill getThinForwardDiagonals();

    ForegroundFill getSquares();

    ForegroundFill getDiamonds();
}
