package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.BorderStyle;
import org.modelcatalogue.spreadsheet.api.Color;
import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.builder.api.BorderDefinition;

public abstract class AbstractBorderDefinition implements BorderDefinition {

    @Override
    public final BorderDefinition style(BorderStyle style) {
        borderStyle = style;
        return this;
    }

    @Override
    public final BorderDefinition color(Color colorPreset) {
        color(colorPreset.getHex());
        return this;
    }

    protected abstract void applyTo(Keywords.BorderSide location);

    protected BorderStyle borderStyle;
}
