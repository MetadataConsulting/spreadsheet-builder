package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.*;

public interface CellStyleDefinition extends ForegroundFillProvider, BorderPositionProvider {

    CellStyleDefinition base(String stylename);

    CellStyleDefinition background(String hexColor);
    CellStyleDefinition background(Color color);

    CellStyleDefinition foreground(String hexColor);
    CellStyleDefinition foreground(Color color);

    CellStyleDefinition fill(ForegroundFill fill);

    CellStyleDefinition font(Configurer<FontDefinition> fontConfiguration);

    /**
     * Sets the indent of the cell in spaces.
     * @param indent the indent of the cell in spaces
     */
    CellStyleDefinition indent(int indent);

    /**
     * Enables word wrapping
     *
     * @param text keyword
     */
    CellStyleDefinition wrap(Keywords.Text text);

    /**
     * Sets the rotation from 0 to 180 (flipped).
     * @param rotation the rotation from 0 to 180 (flipped)
     */
    CellStyleDefinition rotation(int rotation);

    CellStyleDefinition format(String format);

    HorizontalAlignmentConfigurer align(Keywords.VerticalAlignment alignment);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide location, Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderDefinition> borderConfiguration);

}
