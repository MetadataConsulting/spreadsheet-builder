package org.modelcatalogue.builder.spreadsheet.api;

public interface BorderSide {
    BorderSide TOP = PureBorderSide.TOP;
    BorderSide BOTTOM = PureBorderSide.BOTTOM;
    BorderSide LEFT = BorderSideAndHorizontalAlignment.LEFT;
    BorderSide RIGHT = BorderSideAndHorizontalAlignment.RIGHT;

    BorderSide[] BORDER_SIDES = { TOP, BOTTOM, LEFT, RIGHT };
}
