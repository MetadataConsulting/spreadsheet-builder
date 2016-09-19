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

    // TODO: methods with conditions

    void background(String hexColor);
    void background(Color color);

    void foreground(String hexColor);
    void foreground(Color color);

    void fill(ForegroundFill fill);

    void indent(int indent);
    void rotation(int rotation);
    void format(String format);


    // void border()

    void font(@DelegatesTo(FontCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.FontCriterion") Closure fontCriterion);

}
