package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.*;
import builders.dsl.spreadsheet.builder.api.SheetDefinition;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

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

    CellStyleCriterion font(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = FontCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.FontCriterion") Configurer<FontCriterion> fontCriterion);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderCriterion") Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide location, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderCriterion") Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderCriterion") Configurer<BorderCriterion> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleCriterion border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = BorderCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.BorderCriterion") Configurer<BorderCriterion> borderConfiguration);

    CellStyleCriterion having(Predicate<CellStyle> cellStylePredicate);

}
