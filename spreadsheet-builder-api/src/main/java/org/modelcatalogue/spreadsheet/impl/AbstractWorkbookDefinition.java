package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition;
import org.modelcatalogue.spreadsheet.builder.api.Resolvable;
import org.modelcatalogue.spreadsheet.builder.api.Sealable;
import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition;
import org.modelcatalogue.spreadsheet.builder.api.Stylesheet;
import org.modelcatalogue.spreadsheet.builder.api.WorkbookDefinition;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public abstract class AbstractWorkbookDefinition implements WorkbookDefinition {

    protected final Map<String, Configurer<CellStyleDefinition>> namedStylesDefinition = new LinkedHashMap<String, Configurer<CellStyleDefinition>>();
    protected final Map<String, CellStyleDefinition> namedStyles = new LinkedHashMap<String, CellStyleDefinition>();
    protected final Map<String, SheetDefinition> sheets = new LinkedHashMap<String, SheetDefinition>();
    protected final List<Resolvable> toBeResolved = new ArrayList<Resolvable>();

    @Override
    public final WorkbookDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition) {
        namedStylesDefinition.put(name, styleDefinition);
        return this;
    }

    @Override
    public final WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet) {
        try {
            apply(stylesheet.newInstance());
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return this;
    }

    @Override
    public final WorkbookDefinition apply(Stylesheet stylesheet) {
        stylesheet.declareStyles(this);
        return this;
    }

    protected final void resolve() {
        for (Resolvable resolvable : toBeResolved) {
            resolvable.resolve();
        }
    }

    protected abstract CellStyleDefinition createCellStyle();
    protected abstract SheetDefinition createSheet(String name);

    @Override
    public final WorkbookDefinition sheet(String name, Configurer<SheetDefinition> sheetDefinition) {
        SheetDefinition sheet = sheets.get(name);

        if (sheet == null) {
            sheet = createSheet(name);
            sheets.put(name, sheet);
        }

        Configurer.Runner.doConfigure(sheetDefinition, sheet);


        if (sheet instanceof Resolvable) {
            ((Resolvable) sheet).resolve();
        }

        return this;
    }

    protected final CellStyleDefinition getStyle(String name) {
        CellStyleDefinition style = namedStyles.get(name);

        if (style != null) {
            return style;
        }

        style = createCellStyle();
        Configurer.Runner.doConfigure(getStyleDefinition(name), style);

        if (style instanceof Sealable) {
            ((Sealable) style).seal();
        }

        namedStyles.put(name, style);

        return style;
    }

    protected final CellStyleDefinition getStyles(Iterable<String> names) {
        String name = join(names, ".");

        CellStyleDefinition style = namedStyles.get(name);

        if (style != null) {
            return style;
        }

        style = createCellStyle();
        for (String n : names) {
            Configurer.Runner.doConfigure(getStyleDefinition(n), style);
        }

        if (style instanceof Sealable) {
            ((Sealable) style).seal();
        }

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
