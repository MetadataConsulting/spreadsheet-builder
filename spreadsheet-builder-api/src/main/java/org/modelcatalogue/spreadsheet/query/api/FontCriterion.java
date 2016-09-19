package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.builder.api.FontStylesProvider;

public interface FontCriterion extends FontStylesProvider {

    // TODO: methods with conditions

    void color(String hexColor);
    void color(Color color);

    void size(int size);
    void name(String name);

    void make(FontStyle first, FontStyle... other);

}
