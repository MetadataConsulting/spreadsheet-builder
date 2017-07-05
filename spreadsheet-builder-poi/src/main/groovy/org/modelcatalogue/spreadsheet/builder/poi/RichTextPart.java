package org.modelcatalogue.spreadsheet.builder.poi;

class RichTextPart {

    RichTextPart(String text, PoiFontDefinition font, int start, int end) {
        this.text = text;
        this.font = font;
        this.start = start;
        this.end = end;
    }

    public final String getText() {
        return text;
    }

    public final PoiFontDefinition getFont() {
        return font;
    }

    public final int getStart() {
        return start;
    }

    public final int getEnd() {
        return end;
    }

    private final String text;
    private final PoiFontDefinition font;
    private final int start;
    private final int end;
}
