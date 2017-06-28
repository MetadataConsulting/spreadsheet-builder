package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.builder.api.*;

import java.util.*;

public abstract class AbstractWorkbookDefinition<CellStyleType extends CellStyleDefinition & Sealable, SheetType extends AbstractSheetDefinition, SelfType extends WorkbookDefinition> implements WorkbookDefinition {

    protected final Map<String, Configurer<CellStyleDefinition>> namedStylesDefinition = new LinkedHashMap<String, Configurer<CellStyleDefinition>>();
    protected final Map<String, CellStyleType> namedStyles = new LinkedHashMap<String, CellStyleType>();
    protected final Map<String, SheetType> sheets = new LinkedHashMap<String, SheetType>();
    protected final List<Resolvable> toBeResolved = new ArrayList<Resolvable>();

    @Override
    public final SelfType style(String name, Configurer<CellStyleDefinition> styleDefinition) {
        namedStylesDefinition.put(name, styleDefinition);
        return (SelfType) this;
    }

    @Override
    public final SelfType apply(Class<? extends Stylesheet> stylesheet) {
        try {
            apply(stylesheet.newInstance());
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return (SelfType) this;
    }

    @Override
    public final SelfType apply(Stylesheet stylesheet) {
        stylesheet.declareStyles(this);
        return (SelfType) this;
    }

    protected final void resolve() {
        for (Resolvable resolvable : toBeResolved) {
            resolvable.resolve();
        }
    }

    protected abstract CellStyleType createCellStyle();
    protected abstract SheetType createSheet(String name);

    @Override
    public final SelfType sheet(String name, Configurer<SheetDefinition> sheetDefinition) {
        SheetType sheet = sheets.get(name);

        if (sheet == null) {
            sheet = createSheet(name);
            sheets.put(name, sheet);
        }

        Configurer.Runner.doConfigure(sheetDefinition, sheet);

        sheet.processAutoColumns();
        sheet.processAutomaticFilter();
        return (SelfType) this;
    }

    protected final CellStyleType getStyle(String name) {
        CellStyleType style = namedStyles.get(name);

        if (style != null) {
            return style;
        }

        style = createCellStyle();
        Configurer.Runner.doConfigure(getStyleDefinition(name), style);
        style.seal();

        namedStyles.put(name, style);

        return style;
    }

    protected final CellStyleType getStyles(Iterable<String> names) {
        String name = join(names, ".");

        CellStyleType style = namedStyles.get(name);

        if (style != null) {
            return style;
        }

        style = createCellStyle();
        for (String n : names) {
            Configurer.Runner.doConfigure(getStyleDefinition(n), style);
        }
        style.seal();

        namedStyles.put(name, style);

        return style;
    }

    protected final Configurer<CellStyleDefinition> getStyleDefinition(String name) {
        Configurer<CellStyleDefinition> style = namedStylesDefinition.get(name);
        if (style == null) {
            throw new IllegalArgumentException("Style '" + name + "' is not defined");
        }
        return style;
    }


    // from DGM
    private static String join(Iterable<String> array, String separator) {
        StringBuilder buffer = new StringBuilder();
        boolean first = true;

        if (separator == null) separator = "";

        for (String value : array) {
            if (first) {
                first = false;
            } else {
                buffer.append(separator);
            }
            buffer.append(value);
        }
        return buffer.toString();
    }
}
