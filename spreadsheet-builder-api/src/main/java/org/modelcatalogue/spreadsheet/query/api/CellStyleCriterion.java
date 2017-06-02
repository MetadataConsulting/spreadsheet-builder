package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

public interface CellStyleCriterion extends HTMLColorProvider, ForegroundFillProvider {

    void background(String hexColor);
    void background(Color color);
    void background(Predicate<Color> predicate);

    void foreground(String hexColor);
    void foreground(Color color);
    void foreground(Predicate<Color> predicate);

    void fill(ForegroundFill fill);
    void fill(Predicate<ForegroundFill> predicate);

    void indent(int indent);
    void indent(Predicate<Integer> predicate);

    void rotation(int rotation);
    void rotation(Predicate<Integer> predicate);

    void format(String format);
    void format(Predicate<String> format);

    void font(Configurer<FontCriterion> fontCriterion);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    void border(Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    void border(Keywords.BorderSide location, Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderCriterion> borderConfiguration);

    // keywords
    Keywords.PureBorderSide getLeft();
    Keywords.PureBorderSide getRight();

    Keywords.BorderSideAndVerticalAlignment getTop();
    Keywords.BorderSideAndVerticalAlignment getBottom();

}
