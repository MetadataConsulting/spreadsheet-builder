package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.*;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface CellStyleDefinition extends ForegroundFillProvider, BorderPositionProvider, ColorProvider {

    CellStyleDefinition base(String stylename);

    CellStyleDefinition background(String hexColor);
    CellStyleDefinition background(Color color);

    CellStyleDefinition foreground(String hexColor);
    CellStyleDefinition foreground(Color color);

    CellStyleDefinition fill(ForegroundFill fill);

    CellStyleDefinition font(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = FontDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.FontDefinition") Configurer<FontDefinition> fontConfiguration);

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

    CellStyleDefinition align(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderDefinition") Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide location, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderDefinition") Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderDefinition") Configurer<BorderDefinition> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderDefinition") Configurer<BorderDefinition> borderConfiguration);

}
