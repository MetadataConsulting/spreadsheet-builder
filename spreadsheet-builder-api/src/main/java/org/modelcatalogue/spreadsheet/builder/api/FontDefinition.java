package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.api.FontStylesProvider;
import org.modelcatalogue.spreadsheet.api.HTMLColorProvider;

public interface FontDefinition extends HTMLColorProvider, FontStylesProvider {

    FontDefinition color(String hexColor);
    FontDefinition color(Color color);

    FontDefinition size(int size);
    FontDefinition name(String name);

    FontDefinition make(FontStyle first, FontStyle... other);

}
