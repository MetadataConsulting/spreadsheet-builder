package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.BorderStyle;
import builders.dsl.spreadsheet.api.BorderStyleProvider;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.ColorProvider;

public interface BorderDefinition extends BorderStyleProvider, ColorProvider {

    BorderDefinition style(BorderStyle style);
    BorderDefinition color(String hexColor);
    BorderDefinition color(Color color);

}
