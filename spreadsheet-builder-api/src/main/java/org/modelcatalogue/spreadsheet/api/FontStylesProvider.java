package org.modelcatalogue.spreadsheet.api;

import org.modelcatalogue.spreadsheet.api.FontStyle;

public interface FontStylesProvider {

    FontStyle getItalic();
    FontStyle getBold();
    FontStyle getStrikeout();
    FontStyle getUnderline();

}
