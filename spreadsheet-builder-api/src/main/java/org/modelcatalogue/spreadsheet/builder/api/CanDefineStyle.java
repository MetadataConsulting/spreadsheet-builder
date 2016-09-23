package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface CanDefineStyle {
    /**
     * Declare a named style.
     * @param name name of the style
     * @param styleDefinition definition of the style
     */
    void style(String name, @DelegatesTo(CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition") Closure styleDefinition);

    void apply(Class<Stylesheet> stylesheet);
    void apply(Stylesheet stylesheet);
}
