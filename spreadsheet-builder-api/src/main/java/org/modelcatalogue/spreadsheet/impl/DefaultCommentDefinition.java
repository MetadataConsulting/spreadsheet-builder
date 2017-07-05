package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.Comment;
import org.modelcatalogue.spreadsheet.builder.api.CommentDefinition;

public final class DefaultCommentDefinition implements CommentDefinition, Comment {

    @Override
    public CommentDefinition author(String author) {
        this.author = author;
        return this;
    }

    @Override
    public CommentDefinition text(String text) {
        if (text == null) {
            throw new IllegalArgumentException("Text cannot be null");
        }
        this.text = text;
        return this;
    }

    public void width(int widthInCells) {
        assert widthInCells > 0;
        this.width = widthInCells;
    }

    public void height(int heightInCells) {
        assert heightInCells > 0;
        this.height = heightInCells;
    }


    @Override
    public String getAuthor() {
        return author;
    }

    @Override
    public String getText() {
        return text;
    }

    public int getWidth() {
        return width;
    }

    public int getHeight() {
        return height;
    }

    private String author;
    private String text;
    private int width = DEFAULT_WIDTH;
    private int height = DEFAULT_HEIGHT;
}
