package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.*;

import java.util.EnumSet;

public interface FontCriterion extends FontStylesProvider, ColorProvider {

    FontCriterion color(String hexColor);
    FontCriterion color(Color color);
    FontCriterion color(Predicate<Color> predicate);

    FontCriterion size(int size);
    FontCriterion size(Predicate<Integer> predicate);

    FontCriterion name(String name);
    FontCriterion name(Predicate<String> predicate);

    FontCriterion make(FontStyle first, FontStyle... other);
    FontCriterion make(Predicate<EnumSet<FontStyle>> predicate);

    FontCriterion having(Predicate<Font> fontPredicate);
}
