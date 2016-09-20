package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.builder.api.FontStylesProvider;

import java.util.EnumSet;

public interface FontCriterion extends FontStylesProvider {

    void color(String hexColor);
    void color(Color color);
    void color(Condition<Color> condition);

    void size(int size);
    void size(Condition<Integer> condition);

    void name(String name);
    void name(Condition<String> condition);

    void make(FontStyle first, FontStyle... other);
    void make(Condition<EnumSet<FontStyle>> condition);
}
