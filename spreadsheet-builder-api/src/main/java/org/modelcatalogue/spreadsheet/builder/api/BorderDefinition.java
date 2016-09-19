package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.Color;

public interface BorderDefinition {

    void style(BorderStyle style);
    void color(String hexColor);
    void color(Color color);


    // keywords
    BorderStyle getNone();
    BorderStyle getThin();
    BorderStyle getMedium();
    BorderStyle getDashed();
    BorderStyle getDotted();
    BorderStyle getThick();
    BorderStyle getDouble();
    BorderStyle getHair();
    BorderStyle getMediumDashed();
    BorderStyle getDashDot();
    BorderStyle getMediumDashDot();
    BorderStyle getDashDotDot();
    BorderStyle getMediumDashDotDot();
    BorderStyle getSlantedDashDot();

}
