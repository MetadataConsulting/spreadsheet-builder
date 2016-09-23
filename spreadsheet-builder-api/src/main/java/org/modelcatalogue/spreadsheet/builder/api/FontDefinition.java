package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.api.FontStylesProvider;
import org.modelcatalogue.spreadsheet.api.HTMLColorProvider;

public interface FontDefinition extends HTMLColorProvider, FontStylesProvider {

    void color(String hexColor);
    void color(Color color);
    
    void size(int size);
    void name(String name);

    void make(FontStyle first, FontStyle... other);

}
