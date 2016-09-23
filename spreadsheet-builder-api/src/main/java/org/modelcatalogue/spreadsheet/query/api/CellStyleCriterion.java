package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.*;

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

    void font(@DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    void border(@DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide location, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.BorderCriterion") Closure borderConfiguration);

    // keywords
    Keywords.PureBorderSide getLeft();
    Keywords.PureBorderSide getRight();

    Keywords.BorderSideAndVerticalAlignment getTop();
    Keywords.BorderSideAndVerticalAlignment getBottom();

}
