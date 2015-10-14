package org.modelcatalogue.builder.spreadsheet.poi

class RichTextPart {

    final String text
    final PoiFont font
    final int start
    final int end

    RichTextPart(String text, PoiFont font, int start, int end) {
        this.text = text
        this.font = font
        this.start = start
        this.end = end
    }
}
