package org.modelcatalogue.builder.spreadsheet.api;

public interface Font {

    void color(String hexColor);
    void color(Color color);
    
    void size(int size);

    Object getItalic();
    Object getBold();
    Object getStrikeout();
    Object getUnderline();

}
