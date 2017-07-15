package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ColorProvider;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.api.FontStylesProvider;

public interface FontDefinition extends FontStylesProvider, ColorProvider {

    FontDefinition color(String hexColor);
    FontDefinition color(Color color);

    FontDefinition size(int size);
    FontDefinition name(String name);

    FontDefinition style(FontStyle first, FontStyle... other);

}
