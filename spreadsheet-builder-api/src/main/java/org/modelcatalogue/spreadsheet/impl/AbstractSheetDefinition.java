package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.builder.api.*;

import java.util.*;

public abstract class AbstractSheetDefinition implements SheetDefinition, Resolvable {

    private final AbstractWorkbookDefinition workbook;

    private final List<Integer> startPositions = new ArrayList<Integer>();
    private int nextRowNumber = 0;
    protected final Set<Integer> autoColumns = new HashSet<Integer>();
    private final Map<Integer, AbstractRowDefinition> rows = new LinkedHashMap<Integer, AbstractRowDefinition>();
    protected boolean automaticFilter;

    protected AbstractSheetDefinition(AbstractWorkbookDefinition workbook) {
        this.workbook = workbook;
    }

    protected AbstractWorkbookDefinition getWorkbook() {
        return workbook;
    }

    @Override
    public final SheetDefinition row() {
        findOrCreateRow(nextRowNumber++);
        return this;
    }

    private RowDefinition findOrCreateRow(int zeroBasedRowNumber) {
        AbstractRowDefinition row = rows.get(zeroBasedRowNumber + 1);

        if (row != null) {
            return row;
        }

        row = createRow(zeroBasedRowNumber);
        rows.put(zeroBasedRowNumber + 1, row);
        return row;
    }

    protected abstract AbstractRowDefinition createRow(int zeroBasedRowNumber);

    @Override
    public final SheetDefinition row(Configurer<RowDefinition> rowDefinition) {
        RowDefinition row = findOrCreateRow(nextRowNumber++);
        Configurer.Runner.doConfigure(rowDefinition, row);
        return this;
    }

    @Override
    public final SheetDefinition row(int oneBasedRowNumber, Configurer<RowDefinition> rowDefinition) {
        if (oneBasedRowNumber <= 0) {
            throw new IllegalArgumentException("Row index is based on 1. Got: " + oneBasedRowNumber);
        }
        nextRowNumber = oneBasedRowNumber;

        RowDefinition poiRow = findOrCreateRow(oneBasedRowNumber - 1);
        Configurer.Runner.doConfigure(rowDefinition, poiRow);
        return this;
    }

    @Override
    public final SheetDefinition freeze(String column, int row) {
        return freeze(Utils.parseColumn(column), row);
    }

    @Override
    public final SheetDefinition freeze(int column, int row) {
        doFreeze(column, row);
        return this;
    }

    @Override
    public final SheetDefinition lock() {
        doLock();
        return this;
    }

    protected abstract void doLock();

    @Override
    public final SheetDefinition password(String password) {
        doPassword(password);
        return this;
    }

    protected abstract void doPassword(String password);
    protected abstract void doFreeze(int column, int row);

    @Override
    public final SheetDefinition collapse(Configurer<SheetDefinition> insideGroupDefinition) {
        createGroup(true, insideGroupDefinition);
        return this;
    }

    @Override
    public final SheetDefinition group(Configurer<SheetDefinition> insideGroupDefinition) {
        createGroup(false, insideGroupDefinition);
        return this;
    }


    @Override
    public final SheetDefinition filter(Keywords.Auto auto) {
        automaticFilter = true;
        return this;
    }

    @Override
    public final SheetDefinition page(Configurer<PageDefinition> pageDefinition) {
        PageDefinition page = createPageDefinition();
        Configurer.Runner.doConfigure(pageDefinition, page);
        return this;
    }

    protected abstract PageDefinition createPageDefinition();

    private void createGroup(boolean collapsed, Configurer<SheetDefinition> insideGroupDefinition) {
        startPositions.add(nextRowNumber);

        Configurer.Runner.doConfigure(insideGroupDefinition, this);

        int startPosition = startPositions.remove(startPositions.size() - 1);

        if (nextRowNumber - startPosition > 1) {
            int endPosition = nextRowNumber - 1;
            applyRowGroup(startPosition, endPosition, collapsed);
        }

    }

    protected abstract void applyRowGroup(int startPosition, int endPosition, boolean collapsed);

    // TODO: make protected again
    public final void addAutoColumn(int i) {
        autoColumns.add(i);
    }

    protected abstract void processAutoColumns();

    protected abstract void processAutomaticFilter();


    @Override
    public String toString() {
        return "Sheet[" + getName() + "]";
    }

    protected abstract String getName();

    @Override
    public void resolve() {
        processAutoColumns();
        processAutomaticFilter();
    }
}
