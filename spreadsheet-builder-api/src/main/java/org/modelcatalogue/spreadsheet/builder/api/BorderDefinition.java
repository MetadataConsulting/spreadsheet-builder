package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.BorderStyleProvider;
import org.modelcatalogue.spreadsheet.api.Color;

public interface BorderDefinition extends BorderStyleProvider {

    BorderDefinition style(BorderStyle style);
    BorderDefinition color(String hexColor);
    BorderDefinition color(Color color);

}
