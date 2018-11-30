package builders.dsl.spreadsheet.builder.api;

import java.util.function.Consumer;

public interface CanDefineStyle {
    /**
     * Declare a named style.
     * @param name name of the style
     * @param styleDefinition definition of the style
     */
    CanDefineStyle style(String name, Consumer<CellStyleDefinition> styleDefinition);

    CanDefineStyle apply(Class<? extends Stylesheet> stylesheet);
    CanDefineStyle apply(Stylesheet stylesheet);
}
