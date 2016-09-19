package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.BorderStyle;

public abstract class AbstractBorderDefinition implements BorderDefinition {

    @Override
    public BorderStyle getNone() {
        return BorderStyle.NONE;
    }

    @Override
    public BorderStyle getThin() {
        return BorderStyle.THIN;
    }

    @Override
    public BorderStyle getMedium() {
        return BorderStyle.MEDIUM;
    }

    @Override
    public BorderStyle getDashed() {
        return BorderStyle.DASHED;
    }

    @Override
    public BorderStyle getDotted() {
        return BorderStyle.DOTTED;
    }

    @Override
    public BorderStyle getThick() {
        return BorderStyle.THICK;
    }

    @Override
    public BorderStyle getDouble() {
        return BorderStyle.DOUBLE;
    }

    @Override
    public BorderStyle getHair() {
        return BorderStyle.HAIR;
    }

    @Override
    public BorderStyle getMediumDashed() {
        return BorderStyle.MEDIUM_DASHED;
    }

    @Override
    public BorderStyle getDashDot() {
        return BorderStyle.DASH_DOT;
    }

    @Override
    public BorderStyle getMediumDashDot() {
        return BorderStyle.MEDIUM_DASH_DOT;
    }

    @Override
    public BorderStyle getDashDotDot() {
        return BorderStyle.DASH_DOT_DOT;
    }

    @Override
    public BorderStyle getMediumDashDotDot() {
        return BorderStyle.MEDIUM_DASH_DOT_DOT;
    }

    @Override
    public BorderStyle getSlantedDashDot() {
        return BorderStyle.SLANTED_DASH_DOT;
    }

}
