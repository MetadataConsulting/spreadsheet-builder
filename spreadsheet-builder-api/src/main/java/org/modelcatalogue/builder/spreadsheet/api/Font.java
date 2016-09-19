package org.modelcatalogue.builder.spreadsheet.api;

public interface Font extends ProvidesHTMLColors, ProvidesFontStyle {

    void color(String hexColor);
    void color(Color color);
    
    void size(int size);
    void name(String name);

    void make(FontStyle first, FontStyle... other);

}
