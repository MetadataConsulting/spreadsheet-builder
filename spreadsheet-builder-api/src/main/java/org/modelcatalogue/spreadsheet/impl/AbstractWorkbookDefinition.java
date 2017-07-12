package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.builder.api.*;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public abstract class AbstractWorkbookDefinition implements WorkbookDefinition {

    private final Map<String, Configurer<CellStyleDefinition>> namedStylesDefinition = new LinkedHashMap<String, Configurer<CellStyleDefinition>>();
    private final Map<String, AbstractCellStyleDefinition> namedStyles = new LinkedHashMap<String, AbstractCellStyleDefinition>();
    private final Map<String, AbstractSheetDefinition> sheets = new LinkedHashMap<String, AbstractSheetDefinition>();
    private final List<Resolvable> toBeResolved = new ArrayList<Resolvable>();

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

    // TODO: make package private again
    public final void resolve() {
        for (Resolvable resolvable : toBeResolved) {
            resolvable.resolve();
        }
    }

    protected abstract AbstractCellStyleDefinition createCellStyle();
    protected abstract AbstractSheetDefinition createSheet(String name);

    @Override
    public final WorkbookDefinition sheet(String name, Configurer<SheetDefinition> sheetDefinition) {
        AbstractSheetDefinition sheet = sheets.get(name);

        if (sheet == null) {
            sheet = createSheet(name);
            sheets.put(name, sheet);
        }

        Configurer.Runner.doConfigure(sheetDefinition, sheet);

        sheet.resolve();

        return this;
    }

    final AbstractCellStyleDefinition getStyles(Iterable<String> names) {
        String name = Utils.join(names, ".");

        AbstractCellStyleDefinition style = namedStyles.get(name);

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

    final Configurer<CellStyleDefinition> getStyleDefinition(String name) {
        Configurer<CellStyleDefinition> style = namedStylesDefinition.get(name);
        if (style == null) {
            throw new IllegalArgumentException("Style '" + name + "' is not defined");
        }
        return style;
    }

    void addPendingFormula(AbstractPendingFormula formula) {
        toBeResolved.add(formula);
    }

    protected void addPendingLink(AbstractPendingLink link) {
        toBeResolved.add(link);
    }


}
