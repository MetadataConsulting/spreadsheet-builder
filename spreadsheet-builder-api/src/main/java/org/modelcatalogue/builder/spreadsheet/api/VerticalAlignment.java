package org.modelcatalogue.builder.spreadsheet.api;

public interface VerticalAlignment {

    VerticalAlignment TOP = BorderSideAndVerticalAlignment.TOP;
    VerticalAlignment CENTER = PureVerticalAlignment.CENTER;
    VerticalAlignment BOTTOM = BorderSideAndVerticalAlignment.BOTTOM;
    VerticalAlignment JUSTIFY = PureVerticalAlignment.JUSTIFY;
    VerticalAlignment DISTRIBUTED = PureVerticalAlignment.DISTRIBUTED;

    VerticalAlignment[] VERTICAL_ALIGNMENTS = {TOP, CENTER, BOTTOM, JUSTIFY, DISTRIBUTED};

}
