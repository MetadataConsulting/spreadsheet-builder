package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.api.BorderStyle;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.builder.api.BorderDefinition;

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
