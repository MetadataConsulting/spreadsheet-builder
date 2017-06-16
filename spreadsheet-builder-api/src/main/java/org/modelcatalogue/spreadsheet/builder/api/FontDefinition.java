package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.api.FontStylesProvider;

public interface FontDefinition extends FontStylesProvider {

    FontDefinition color(String hexColor);
    FontDefinition color(Color color);

    FontDefinition size(int size);
    FontDefinition name(String name);

    FontDefinition make(FontStyle first, FontStyle... other);

}
