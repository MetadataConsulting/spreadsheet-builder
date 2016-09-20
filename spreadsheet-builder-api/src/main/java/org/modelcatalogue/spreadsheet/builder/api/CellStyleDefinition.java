package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.*;

public interface CellStyleDefinition extends HTMLColorProvider, ForegroundFillProvider {

    void base(String stylename);

    void background(String hexColor);
    void background(Color color);

    void foreground(String hexColor);
    void foreground(Color color);

    void fill(ForegroundFill fill);

    void font(@DelegatesTo(FontDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.FontDefinition") Closure fontConfiguration);

    /**
     * Sets the indent of the cell in spaces.
     * @param indent the indent of the cell in spaces
     */
    void indent(int indent);

    /**
     * Enables word wrapping
     *
     * @param text keyword
     */
    void wrap(Keywords.Text text);

    Keywords.Text getText();

    /**
     * Sets the rotation from 0 to 180 (flipped).
     * @param rotation the rotation from 0 to 180 (flipped)
     */
    void rotation(int rotation);

    void format(String format);

    HorizontalAlignmentConfigurer align(Keywords.VerticalAlignment alignment);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration closure
     */
    void border(@DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide location, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration closure
     */
    void border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(BorderDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.BorderDefinition") Closure borderConfiguration);

    Keywords.PureVerticalAlignment getCenter();
    Keywords.PureVerticalAlignment getJustify();
    Keywords.PureVerticalAlignment getDistributed();

    // keywords
    Keywords.PureBorderSide getLeft();
    Keywords.PureBorderSide getRight();

    Keywords.BorderSideAndVerticalAlignment getTop();
    Keywords.BorderSideAndVerticalAlignment getBottom();
}
