package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;
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
    CanDefineStyle style(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);

    CanDefineStyle apply(Class<? extends Stylesheet> stylesheet);
    CanDefineStyle apply(Stylesheet stylesheet);
}
