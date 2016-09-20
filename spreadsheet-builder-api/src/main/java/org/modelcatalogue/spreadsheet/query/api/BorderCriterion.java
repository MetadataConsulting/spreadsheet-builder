package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.BorderStyleProvider;
import org.modelcatalogue.spreadsheet.api.Color;

public interface BorderCriterion extends BorderStyleProvider {

    void style(BorderStyle style);
    void style(Predicate<BorderStyle> predicate);

    void color(String hexColor);
    void color(Color color);
    void color(Predicate<Color> predicate);

}
