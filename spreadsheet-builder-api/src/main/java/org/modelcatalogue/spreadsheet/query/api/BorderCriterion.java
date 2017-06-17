package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Border;
import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.BorderStyleProvider;
import org.modelcatalogue.spreadsheet.api.Color;

public interface BorderCriterion extends BorderStyleProvider {

    BorderCriterion style(BorderStyle style);
    BorderCriterion style(Predicate<BorderStyle> predicate);

    BorderCriterion color(String hexColor);
    BorderCriterion color(Color color);
    BorderCriterion color(Predicate<Color> predicate);

    BorderCriterion having(Predicate<Border> borderPredicate);

}
