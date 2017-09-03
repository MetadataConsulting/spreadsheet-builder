package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Border;
import builders.dsl.spreadsheet.api.BorderStyle;
import builders.dsl.spreadsheet.api.BorderStyleProvider;
import builders.dsl.spreadsheet.api.Color;

public interface BorderCriterion extends BorderStyleProvider {

    BorderCriterion style(BorderStyle style);
    BorderCriterion style(Predicate<BorderStyle> predicate);

    BorderCriterion color(String hexColor);
    BorderCriterion color(Color color);
    BorderCriterion color(Predicate<Color> predicate);

    BorderCriterion having(Predicate<Border> borderPredicate);

}
