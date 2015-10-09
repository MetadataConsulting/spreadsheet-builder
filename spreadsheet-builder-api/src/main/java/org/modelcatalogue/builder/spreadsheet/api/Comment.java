package org.modelcatalogue.builder.spreadsheet.api;

public interface Comment {

    int DEFAULT_WIDTH = 3;
    int DEFAULT_HEIGHT = 3;

    void author(String author);
    void text(String text);

}
