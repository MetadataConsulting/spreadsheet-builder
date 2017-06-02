package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.FontStyle;
import org.modelcatalogue.spreadsheet.api.FontStylesProvider;

import java.util.EnumSet;

public interface FontCriterion extends FontStylesProvider {

    FontCriterion color(String hexColor);
    FontCriterion color(Color color);
    FontCriterion color(Predicate<Color> predicate);

    FontCriterion size(int size);
    FontCriterion size(Predicate<Integer> predicate);

    FontCriterion name(String name);
    FontCriterion name(Predicate<String> predicate);

    FontCriterion make(FontStyle first, FontStyle... other);
    FontCriterion make(Predicate<EnumSet<FontStyle>> predicate);
}
