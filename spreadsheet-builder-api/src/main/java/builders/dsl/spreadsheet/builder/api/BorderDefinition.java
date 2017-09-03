package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.BorderStyle;
import builders.dsl.spreadsheet.api.BorderStyleProvider;
import builders.dsl.spreadsheet.api.Color;

public interface BorderDefinition extends BorderStyleProvider {

    BorderDefinition style(BorderStyle style);
    BorderDefinition color(String hexColor);
    BorderDefinition color(Color color);

}
