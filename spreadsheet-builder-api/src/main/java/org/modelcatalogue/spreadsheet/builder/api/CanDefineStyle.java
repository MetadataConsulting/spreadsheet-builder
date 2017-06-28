package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Configurer;

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
