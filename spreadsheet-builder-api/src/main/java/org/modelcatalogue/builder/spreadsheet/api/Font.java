package org.modelcatalogue.builder.spreadsheet.api;

public interface Font extends ProvidesHTMLColors {

    void color(String hexColor);
    void color(Color color);
    
    void size(int size);
    void name(String name);

    Object getItalic();
    Object getBold();
    Object getStrikeout();
    Object getUnderline();

}
