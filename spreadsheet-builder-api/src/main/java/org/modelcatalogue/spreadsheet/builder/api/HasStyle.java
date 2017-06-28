package org.modelcatalogue.spreadsheet.builder.api;


import org.modelcatalogue.spreadsheet.api.Configurer;

public interface HasStyle {

    /**
     * Applies a customized named style to the current element.
     *
     * @param name the name of the style
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    HasStyle style(String name, Configurer<CellStyleDefinition> styleDefinition);

    /**
     * Applies a customized named style to the current element.
     *
     * @param names the names of the styles
     * @param styleDefinition the definition of the style customizing the predefined style
     */
    HasStyle styles(Iterable<String> names, Configurer<CellStyleDefinition> styleDefinition);

    /**
     * Applies the style defined by the closure to the current element.
     * @param styleDefinition the definition of the style
     */
    HasStyle style(Configurer<CellStyleDefinition> styleDefinition);

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
