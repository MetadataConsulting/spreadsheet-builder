package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.FontStylesProvider;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.ColorProvider;
import builders.dsl.spreadsheet.api.FontStyle;

public interface FontDefinition extends FontStylesProvider, ColorProvider {

    FontDefinition color(String hexColor);
    FontDefinition color(Color color);

    FontDefinition size(int size);
    FontDefinition name(String name);

    FontDefinition style(FontStyle first, FontStyle... other);

}
