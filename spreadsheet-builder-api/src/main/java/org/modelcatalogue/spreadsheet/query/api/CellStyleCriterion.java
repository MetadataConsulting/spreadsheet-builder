package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.api.ForegroundFills;
import org.modelcatalogue.spreadsheet.api.HTMLColors;

public interface CellStyleCriterion extends HTMLColors, ForegroundFills {

    void background(String hexColor);
    void background(Color color);
    void background(Predicate<Color> predicate);

    void foreground(String hexColor);
    void foreground(Color color);
    void foreground(Predicate<Color> predicate);

    void fill(ForegroundFill fill);
    void fill(Predicate<ForegroundFill> predicate);

    void indent(int indent);
    void indent(Predicate<Integer> predicate);

    void rotation(int rotation);
    void rotation(Predicate<Integer> predicate);

    void format(String format);
    void format(Predicate<String> format);

    // void border()

    void font(@DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion);

}
