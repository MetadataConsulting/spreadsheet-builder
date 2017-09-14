package builders.dsl.spreadsheet.builder.api;


import builders.dsl.spreadsheet.api.Configurer;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface HasStyle {

    /**
     * Applies a customized named style to the current element.
     *
     * @param name the name of the style
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    HasStyle style(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);

    /**
     * Applies a customized named style to the current element.
     *
     * @param names the names of the styles
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    HasStyle styles(Iterable<String> names, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);

    HasStyle styles(Iterable<String> styles, Iterable<Configurer<CellStyleDefinition>> styleDefinitions);

    /**
     * Applies the style defined by the closure to the current element.
     * @param styleDefinition the definition of the style
     */
    HasStyle style(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);

    /**
     * Applies the named style to the current element.
     *
     * The style can be changed no longer.
     *
     * @param name the name of the style
     */
    HasStyle style(String name);

    /**
     * Applies the named style to the current element.
     *
     * The style can be changed no longer.
     *
     * @param names style names to be applied
     */
    HasStyle styles(String... names);
    /**
     * Applies the named style to the current element.
     *
     * The style can be changed no longer.
     *
     * @param names style names to be applied
     */
    HasStyle styles(Iterable<String> names);
}
