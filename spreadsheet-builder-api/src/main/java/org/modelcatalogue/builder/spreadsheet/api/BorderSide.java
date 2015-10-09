package org.modelcatalogue.builder.spreadsheet.api;

public interface BorderSide {
    BorderSide LEFT = PureBorderSide.LEFT;
    BorderSide RIGHT = PureBorderSide.RIGHT;
    BorderSide TOP = BorderSideAndVerticalAlignment.TOP;
    BorderSide BOTTOM = BorderSideAndVerticalAlignment.BOTTOM;

    BorderSide[] BORDER_SIDES = { TOP, BOTTOM, LEFT, RIGHT };
}
