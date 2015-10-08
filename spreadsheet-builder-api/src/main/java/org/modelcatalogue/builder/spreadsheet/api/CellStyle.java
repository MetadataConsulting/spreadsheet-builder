package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;

public interface CellStyle extends ProvidesHTMLColors {

    void background(String hexColor);
    void background(Color color);

    void foreground(String hexColor);
    void foreground(Color color);

    void fill(ForegroundFill fill);

    void font(@DelegatesTo(Font.class) Closure fontConfiguration);

    /**
     * Sets the indent of the cell in spaces.
     * @param indent the indent of the cell in spaces
     */
    void indent(int indent);

    /**
     * Sets the cell to be locked.
     *
     */
    Object getLocked();

    /**
     * Sets the cell to wrap the words.
     *
     */
    Object getWrapped();

    /**
     * Sets the cell to be hidden.
     *
     */
    Object getHidden();

    /**
     * Sets the rotation from 0 to 180 (flipped).
     * @param rotation the rotation from 0 to 180 (flipped)
     */
    void rotation(int rotation);

    void format(String format);

    VerticalAlignmentConfigurer align(HorizontalAlignment alignment);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    void border(@DelegatesTo(Border.class) Closure borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(BorderSide location, @DelegatesTo(Border.class) Closure borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(BorderSide first, BorderSide second, @DelegatesTo(Border.class) Closure borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(BorderSide first, BorderSide second, BorderSide third, @DelegatesTo(Border.class) Closure borderConfiguration);



    // keywords

    BorderSideAndHorizontalAlignment getLeft();
    BorderSideAndHorizontalAlignment getRight();

    BorderSide getTop();
    BorderSide getBottom();

    PureHorizontalAlignment getGeneral();
    PureHorizontalAlignment getCenter();
    PureHorizontalAlignment getFill();
    PureHorizontalAlignment getJustify();
    PureHorizontalAlignment getCenterSelection();

    ForegroundFill getNoFill();
    ForegroundFill getSolidForeground();
    ForegroundFill getFineDots();
    ForegroundFill getAltBars();
    ForegroundFill getSparseDots();
    ForegroundFill getThickHorizontalBands();
    ForegroundFill getThickVerticalBands();
    ForegroundFill getThickBackwardDiagonal();
    ForegroundFill getThickForwardDiagonal();
    ForegroundFill getBigSpots();
    ForegroundFill getBricks();
    ForegroundFill getThinHorizontalBands();
    ForegroundFill getThinVerticalBands();
    ForegroundFill getThinBackwardDiagonal();
    ForegroundFill getThinForwardDiagonal();
    ForegroundFill getSquares();
    ForegroundFill getDiamonds();






}
