package org.modelcatalogue.spreadsheet.builder.api;

public interface CommentDefinition {

    int DEFAULT_WIDTH = 3;
    int DEFAULT_HEIGHT = 3;

    void author(String author);
    void text(String text);

}
