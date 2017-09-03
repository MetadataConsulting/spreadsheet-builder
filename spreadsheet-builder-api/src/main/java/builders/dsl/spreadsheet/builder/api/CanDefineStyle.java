package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;

public interface CanDefineStyle {
    /**
     * Declare a named style.
     * @param name name of the style
     * @param styleDefinition definition of the style
     */
    CanDefineStyle style(String name, Configurer<CellStyleDefinition> styleDefinition);

    CanDefineStyle apply(Class<? extends Stylesheet> stylesheet);
    CanDefineStyle apply(Stylesheet stylesheet);
}
