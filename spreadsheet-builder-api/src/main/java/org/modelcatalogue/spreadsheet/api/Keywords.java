package org.modelcatalogue.spreadsheet.api;

public final class Keywords {


    public static final BorderStyle none = BorderStyle.NONE;
    public static final BorderStyle thin = BorderStyle.THIN;
    public static final BorderStyle medium = BorderStyle.MEDIUM;
    public static final BorderStyle dashed = BorderStyle.DASHED;
    public static final BorderStyle dotted = BorderStyle.DOTTED;
    public static final BorderStyle thick = BorderStyle.THICK;
    public static final BorderStyle doubleBorder = BorderStyle.DOUBLE;
    public static final BorderStyle hair = BorderStyle.HAIR;
    public static final BorderStyle mediumDashed = BorderStyle.MEDIUM_DASHED;
    public static final BorderStyle mashDot = BorderStyle.DASH_DOT;
    public static final BorderStyle mediumDashDot = BorderStyle.MEDIUM_DASH_DOT;
    public static final BorderStyle dashDotDot = BorderStyle.DASH_DOT_DOT;
    public static final BorderStyle mediumDashDotDot = BorderStyle.MEDIUM_DASH_DOT_DOT;
    public static final BorderStyle slantedDashDot = BorderStyle.SLANTED_DASH_DOT;
    public static final Keywords.Orientation portrait = Keywords.Orientation.PORTRAIT;
    public static final Keywords.Orientation landscape = Keywords.Orientation.LANDSCAPE;
    public static final Keywords.Fit width = Keywords.Fit.WIDTH;
    public static final Keywords.Fit height = Keywords.Fit.HEIGHT;
    public static final Keywords.Paper letter = Keywords.Paper.LETTER;
    public static final Keywords.Paper letterSmall = Keywords.Paper.LETTER_SMALL;
    public static final Keywords.Paper tabloid = Keywords.Paper.TABLOID;
    public static final Keywords.Paper ledger = Keywords.Paper.LEDGER;
    public static final Keywords.Paper legal = Keywords.Paper.LEGAL;
    public static final Keywords.Paper statement = Keywords.Paper.STATEMENT;
    public static final Keywords.Paper executive = Keywords.Paper.EXECUTIVE;
    public static final Keywords.Paper A3 = Keywords.Paper.A3;
    public static final Keywords.Paper A4 = Keywords.Paper.A4;
    public static final Keywords.Paper A4Small = Keywords.Paper.A4_SMALL;
    public static final Keywords.Paper A5 = Keywords.Paper.A5;
    public static final Keywords.Paper B4 = Keywords.Paper.B4;
    public static final Keywords.Paper B5 = Keywords.Paper.B5;
    public static final Keywords.Paper folio = Keywords.Paper.FOLIO;
    public static final Keywords.Paper quarto = Keywords.Paper.QUARTO;
    public static final Keywords.Paper standard10x14 = Keywords.Paper.STANDARD_10_14;
    public static final Keywords.Paper standard11x17 = Keywords.Paper.STANDARD_11_17;
    public static final FontStyle italic = FontStyle.ITALIC;
    public static final FontStyle bold = FontStyle.BOLD;
    public static final FontStyle strikeout = FontStyle.STRIKEOUT;
    public static final FontStyle underline = FontStyle.UNDERLINE;
    public static final ForegroundFill noFill = ForegroundFill.NO_FILL;
    public static final ForegroundFill solidForeground = ForegroundFill.SOLID_FOREGROUND;
    public static final ForegroundFill fineDots = ForegroundFill.FINE_DOTS;
    public static final ForegroundFill altBars = ForegroundFill.ALT_BARS;
    public static final ForegroundFill sparseDots = ForegroundFill.SPARSE_DOTS;
    public static final ForegroundFill thickHorizontalBands = ForegroundFill.THICK_HORZ_BANDS;
    public static final ForegroundFill thickVerticalBands = ForegroundFill.THICK_VERT_BANDS;
    public static final ForegroundFill thickBackwardDiagonals = ForegroundFill.THICK_BACKWARD_DIAG;
    public static final ForegroundFill thickForwardDiagonals = ForegroundFill.THICK_FORWARD_DIAG;
    public static final ForegroundFill bigSpots = ForegroundFill.BIG_SPOTS;
    public static final ForegroundFill bricks = ForegroundFill.BRICKS;
    public static final ForegroundFill thinHorizontalBands = ForegroundFill.THIN_HORZ_BANDS;
    public static final ForegroundFill thinVerticalBands = ForegroundFill.THIN_VERT_BANDS;
    public static final ForegroundFill thinBackwardDiagonals = ForegroundFill.THIN_BACKWARD_DIAG;
    public static final ForegroundFill thinForwardDiagonals = ForegroundFill.THICK_FORWARD_DIAG;
    public static final ForegroundFill squares = ForegroundFill.SQUARES;
    public static final ForegroundFill diamonds = ForegroundFill.DIAMONDS;
    public static final Keywords.PureBorderSide left = Keywords.PureBorderSide.LEFT;
    public static final Keywords.PureBorderSide right = Keywords.PureBorderSide.RIGHT;
    public static final Keywords.BorderSideAndVerticalAlignment top = Keywords.BorderSideAndVerticalAlignment.TOP;
    public static final Keywords.BorderSideAndVerticalAlignment bottom = Keywords.BorderSideAndVerticalAlignment.BOTTOM;
    public static final Keywords.PureVerticalAlignment center = Keywords.PureVerticalAlignment.CENTER;
    public static final Keywords.PureVerticalAlignment justify = Keywords.PureVerticalAlignment.JUSTIFY;
    public static final Keywords.PureVerticalAlignment distributed = Keywords.PureVerticalAlignment.DISTRIBUTED;
    public static final Keywords.Text text = Keywords.Text.WRAP;
    public static final Keywords.Auto auto = Keywords.Auto.AUTO;
    public static final Keywords.To to = Keywords.To.TO;
    public static final Keywords.Image image = Keywords.Image.IMAGE;


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
