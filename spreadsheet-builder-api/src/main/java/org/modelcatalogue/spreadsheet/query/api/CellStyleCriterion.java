package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.api.Configurer;

public interface CellStyleCriterion extends ForegroundFillProvider, BorderPositionProvider, ColorProvider {

    CellStyleCriterion background(String hexColor);
    CellStyleCriterion background(Color color);
    CellStyleCriterion background(Predicate<Color> predicate);

    CellStyleCriterion foreground(String hexColor);
    CellStyleCriterion foreground(Color color);
    CellStyleCriterion foreground(Predicate<Color> predicate);

    CellStyleCriterion fill(ForegroundFill fill);
    CellStyleCriterion fill(Predicate<ForegroundFill> predicate);

    CellStyleCriterion indent(int indent);
    CellStyleCriterion indent(Predicate<Integer> predicate);

    CellStyleCriterion rotation(int rotation);
    CellStyleCriterion rotation(Predicate<Integer> predicate);

    CellStyleCriterion format(String format);
    CellStyleCriterion format(Predicate<String> format);

    CellStyleCriterion font(Configurer<FontCriterion> fontCriterion);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide location, Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderCriterion> borderConfiguration);

    CellStyleCriterion having(Predicate<CellStyle> cellStylePredicate);

}
