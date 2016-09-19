package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.ForegroundFill;
import org.modelcatalogue.spreadsheet.api.ForegroundFills;
import org.modelcatalogue.spreadsheet.api.HTMLColors;

public interface CellStyleCriterion extends HTMLColors, ForegroundFills {

    void background(String hexColor);
    void background(Color color);

    void foreground(String hexColor);
    void foreground(Color color);

    void fill(ForegroundFill fill);

    void indent(int indent);
    void rotation(int rotation);
    void format(String format);

}
