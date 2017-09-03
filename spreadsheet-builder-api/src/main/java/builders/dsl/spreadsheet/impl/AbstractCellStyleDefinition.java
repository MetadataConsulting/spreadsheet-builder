package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Color;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.ForegroundFill;
import builders.dsl.spreadsheet.builder.api.BorderDefinition;
import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;
import builders.dsl.spreadsheet.builder.api.FontDefinition;
import builders.dsl.spreadsheet.builder.api.Sealable;

public abstract class AbstractCellStyleDefinition implements CellStyleDefinition, Sealable {

    protected AbstractCellStyleDefinition(AbstractCellDefinition cell) {
        this.workbook = cell.getRow().getSheet().getWorkbook();
    }

    protected AbstractCellStyleDefinition(AbstractWorkbookDefinition workbook) {
        this.workbook = workbook;
    }

    @Override
    public final CellStyleDefinition base(String stylename) {
        checkSealed();
        Configurer.Runner.doConfigure(workbook.getStyleDefinition(stylename), this);
        return this;
    }

    public final void checkSealed() {
        if (sealed) {
            throw new IllegalStateException("The cell style is already sealed! You need to create new style. Use 'styles' method to combine multiple named styles! Create new named style if you're trying to update existing style with closure definition.");
        }

    }

    public final boolean isSealed() {
        return sealed;
    }

    @Override
    public final CellStyleDefinition background(String hexColor) {
        checkSealed();
        doBackground(hexColor);
        return this;
    }

    protected abstract void doBackground(String hexColor);

    @Override
    public final CellStyleDefinition background(Color colorPreset) {
        checkSealed();
        background(colorPreset.getHex());
        return this;
    }

    @Override
    public final CellStyleDefinition foreground(String hexColor) {
        checkSealed();
        doForeground(hexColor);
        return this;
    }

    protected abstract void doForeground(String hexColor);

    @Override
    public final CellStyleDefinition foreground(Color colorPreset) {
        checkSealed();
        foreground(colorPreset.getHex());
        return this;
    }

    @Override
    public final CellStyleDefinition fill(ForegroundFill fill) {
        checkSealed();
        doFill(fill);
        return this;
    }

    protected abstract void doFill(ForegroundFill fill);

    @Override
    public final CellStyleDefinition font(Configurer<FontDefinition> fontConfiguration) {
        checkSealed();
        if (font == null) {
            font = createFont();
        }

        Configurer.Runner.doConfigure(fontConfiguration, font);
        return this;
    }

    protected abstract FontDefinition createFont();

    @Override
    public final CellStyleDefinition indent(int indent) {
        checkSealed();
        doIndent(indent);
        return this;
    }

    protected abstract void doIndent(int indent);

    @Override
    public final CellStyleDefinition wrap(Keywords.Text text) {
        checkSealed();
        doWrapText();
        return this;
    }

    protected abstract void doWrapText();

    @Override
    public final CellStyleDefinition rotation(int rotation) {
        checkSealed();
        doRotation(rotation);
        return this;
    }

    protected abstract void doRotation(int rotation);

    @Override
    public final CellStyleDefinition format(String format) {
        checkSealed();
        doFormat(format);
        return this;
    }

    protected abstract void doFormat(String format);

    @Override
    public final CellStyleDefinition align(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment) {
        checkSealed();
        doAlign(verticalAlignment, horizontalAlignment);
        return this;
    }

    protected abstract void doAlign(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment);

    @Override
    public final CellStyleDefinition border(Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        AbstractBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);

        for (Keywords.BorderSide side: Keywords.BorderSide.BORDER_SIDES) {
            poiBorder.applyTo(side);
        }
        return this;
    }

    @Override
    public final CellStyleDefinition border(Keywords.BorderSide location, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        AbstractBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(location);
        return this;
    }

    @Override
    public final CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        AbstractBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(first);
        poiBorder.applyTo(second);
        return this;
    }

    @Override
    public final CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Configurer<BorderDefinition> borderConfiguration) {
        checkSealed();
        AbstractBorderDefinition poiBorder = findOrCreateBorder();
        Configurer.Runner.doConfigure(borderConfiguration, poiBorder);
        poiBorder.applyTo(first);
        poiBorder.applyTo(second);
        poiBorder.applyTo(third);
        return this;
    }

    private AbstractBorderDefinition findOrCreateBorder() {
        if (border == null) {
            border = createBorder();
        }

        return border;
    }

    protected abstract AbstractBorderDefinition createBorder();

    public final void seal() {
        this.sealed = true;
    }

    protected abstract void assignTo(AbstractCellDefinition cell);

    protected final AbstractWorkbookDefinition workbook;
    private FontDefinition font;
    private AbstractBorderDefinition border;
    private boolean sealed;
}
