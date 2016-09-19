package org.modelcatalogue.spreadsheet.api;

import java.util.EnumSet;

public interface Font {

    Color getColor();
    int getSize();
    String getName();
    EnumSet<FontStyle> getStyles();

}
