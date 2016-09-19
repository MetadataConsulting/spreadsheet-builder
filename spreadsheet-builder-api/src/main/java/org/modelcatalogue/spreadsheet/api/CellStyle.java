package org.modelcatalogue.spreadsheet.api;

public interface CellStyle {

    Color getForeground();
    Color getBackground();
    ForegroundFill getFill();
    int getIndent();
    int getRotation();
    String getFormat();
    Font getFont();

}
