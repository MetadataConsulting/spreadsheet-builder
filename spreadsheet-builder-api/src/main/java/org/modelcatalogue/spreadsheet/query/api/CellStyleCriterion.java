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
    void background(Condition<Color> condition);

    void foreground(String hexColor);
    void foreground(Color color);
    void foreground(Condition<Color> condition);

    void fill(ForegroundFill fill);
    void fill(Condition<ForegroundFill> condition);

    void indent(int indent);
    void indent(Condition<Integer> condition);

    void rotation(int rotation);
    void rotation(Condition<Integer> condition);

    void format(String format);
    void format(Condition<String> format);

    // void border()

    void font(@DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion);

}
