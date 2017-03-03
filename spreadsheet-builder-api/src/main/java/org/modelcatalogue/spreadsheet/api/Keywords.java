package org.modelcatalogue.spreadsheet.api;

public final class Keywords {

    private Keywords() {}

    public enum Text {
        WRAP
    }

    public enum To {
        TO
    }

    public enum Image {
        IMAGE
    }

    public enum Auto {
        AUTO
    }

    public enum BorderSideAndVerticalAlignment implements BorderSide, VerticalAlignment {
        TOP,
        BOTTOM
    }

    public enum PureBorderSide implements BorderSide {
        LEFT,
        RIGHT
    }

    public enum PureVerticalAlignment implements VerticalAlignment {
        CENTER,
        JUSTIFY,
        DISTRIBUTED
    }

    public enum Orientation {
        LANDSCAPE,
        PORTRAIT
    }

    public enum Fit {
        HEIGHT,
        WIDTH
    }

    public enum Paper {
        LETTER,
        LETTER_SMALL,
        TABLOID,
        LEDGER,
        LEGAL,
        STATEMENT,
        EXECUTIVE,
        A3,
        A4,
        A4_SMALL,
        A5,
        B4,
        B5,
        FOLIO,
        QUARTO,
        STANDARD_10_14,
        STANDARD_11_17
    }

    //CHECKSTYLE:OFF
    public interface BorderSide {
        BorderSide LEFT = PureBorderSide.LEFT;
        BorderSide RIGHT = PureBorderSide.RIGHT;
        BorderSide TOP = BorderSideAndVerticalAlignment.TOP;
        BorderSide BOTTOM = BorderSideAndVerticalAlignment.BOTTOM;

        BorderSide[] BORDER_SIDES = {TOP, BOTTOM, LEFT, RIGHT};
    }

    public interface VerticalAlignment {

        VerticalAlignment TOP = BorderSideAndVerticalAlignment.TOP;
        VerticalAlignment CENTER = PureVerticalAlignment.CENTER;
        VerticalAlignment BOTTOM = BorderSideAndVerticalAlignment.BOTTOM;
        VerticalAlignment JUSTIFY = PureVerticalAlignment.JUSTIFY;
        VerticalAlignment DISTRIBUTED = PureVerticalAlignment.DISTRIBUTED;

        VerticalAlignment[] VERTICAL_ALIGNMENTS = {TOP, CENTER, BOTTOM, JUSTIFY, DISTRIBUTED};

    }
    //CHECKSTYLE:ON
}
