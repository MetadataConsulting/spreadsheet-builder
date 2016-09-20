package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.builder.api.FontStylesProvider;

import java.util.EnumSet;

public interface FontCriterion extends FontStylesProvider {

    void color(String hexColor);
    void color(Color color);
    void color(Predicate<Color> predicate);

    void size(int size);
    void size(Predicate<Integer> predicate);

    void name(String name);
    void name(Predicate<String> predicate);

    void make(FontStyle first, FontStyle... other);
    void make(Predicate<EnumSet<FontStyle>> predicate);
}
