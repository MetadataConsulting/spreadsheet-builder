package org.modelcatalogue.spreadsheet.builder.poi

class RichTextPart {

    final String text
    final PoiFontDefinition font
    final int start
    final int end

    RichTextPart(String text, PoiFontDefinition font, int start, int end) {
        this.text = text
        this.font = font
        this.start = start
        this.end = end
    }
}
