package org.modelcatalogue.spreadsheet.api;

public interface BorderStyleProvider {
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
