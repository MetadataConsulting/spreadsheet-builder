package org.modelcatalogue.builder.spreadsheet.api;

public interface HorizontalAlignment {

    HorizontalAlignment GENERAL = PureHorizontalAlignment.GENERAL;
    HorizontalAlignment CENTER = PureHorizontalAlignment.CENTER;
    HorizontalAlignment FILL = PureHorizontalAlignment.FILL;
    HorizontalAlignment JUSTIFY = PureHorizontalAlignment.JUSTIFY;
    HorizontalAlignment CENTER_SELECTION = PureHorizontalAlignment.CENTER_SELECTION;
    HorizontalAlignment LEFT = BorderSideAndHorizontalAlignment.LEFT;
    HorizontalAlignment RIGHT = BorderSideAndHorizontalAlignment.RIGHT;

}
