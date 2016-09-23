package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.BorderStyleProvider;
import org.modelcatalogue.spreadsheet.api.Color;

public interface BorderDefinition extends BorderStyleProvider {

    void style(BorderStyle style);
    void color(String hexColor);
    void color(Color color);

}
