package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.Spannable;
import org.modelcatalogue.spreadsheet.builder.api.*;

import java.util.*;

public abstract class AbstractRowDefinition implements RowDefinition {

    protected final AbstractSheetDefinition sheet;

    private List<String> styles = new ArrayList<String>();
    private List<Configurer<CellStyleDefinition>> styleDefinitions = new ArrayList<Configurer<CellStyleDefinition>>();

    private final List<Integer> startPositions = new ArrayList<Integer>();
    private int nextColNumber = 0;
    private final Map<Integer, CellDefinition> cells = new LinkedHashMap<Integer, CellDefinition>();

    protected AbstractRowDefinition(AbstractSheetDefinition sheet) {
        this.sheet = sheet;
    }

    private CellDefinition findOrCreateCell(int zeroBasedCellNumber) {
        CellDefinition cell = cells.get((zeroBasedCellNumber + 1));

        if (cell != null) {
            return cell;
        }

        cell = createCell(zeroBasedCellNumber);

        cells.put(zeroBasedCellNumber + 1, cell);

        return cell;
    }

    protected abstract CellDefinition createCell(int zeroBasedCellNumber);

    @Override
    public final RowDefinition cell() {
        cell(null);
        return this;
    }

    @Override
    public final RowDefinition cell(Object value) {
        CellDefinition cell = findOrCreateCell(nextColNumber++);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            cell.styles(styles, styleDefinitions);
        }

        cell.value(value);

        if (cell instanceof Resolvable) {
            ((Resolvable) cell).resolve();
        }
        return this;
    }

    @Override
    public final RowDefinition cell(Configurer<CellDefinition> cellDefinition) {
        CellDefinition poiCell = findOrCreateCell(nextColNumber);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            poiCell.styles(styles, styleDefinitions);
        }

        Configurer.Runner.doConfigure(cellDefinition, poiCell);

        if (poiCell instanceof Spannable) {
            nextColNumber += ((Spannable) poiCell).getColspan();
        } else {
            nextColNumber++;
        }

        handleSpans(poiCell);

        if (poiCell instanceof Resolvable) {
            ((Resolvable) poiCell).resolve();
        }

        return this;
    }

    // TODO: change to abstract cell definition later on
    protected abstract void handleSpans(CellDefinition poiCell);

    @Override
    public final RowDefinition cell(int column, Configurer<CellDefinition> cellDefinition) {
        CellDefinition poiCell = findOrCreateCell(column - 1);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            poiCell.styles(styles, styleDefinitions);
        }

        Configurer.Runner.doConfigure(cellDefinition, poiCell);

        if (poiCell instanceof Spannable) {
            nextColNumber = column - 1 + ((Spannable) poiCell).getColspan();
        } else {
            nextColNumber = column;
        }

        handleSpans(poiCell);

        if (poiCell instanceof Resolvable) {
            ((Resolvable) poiCell).resolve();
        }

        return this;
    }

    @Override
    public final RowDefinition cell(String column, Configurer<CellDefinition> cellDefinition) {
        cell(Cell.Util.parseColumn(column), cellDefinition);
        return this;
    }

    @Override
    public final RowDefinition style(Configurer<CellStyleDefinition> styleDefinition) {
        styleDefinitions.add(styleDefinition);
        return this;
    }

    @Override
    public final RowDefinition style(String name) {
        styles.add(name);
        return this;
    }

    @Override
    public final RowDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition) {
        style(name);
        style(styleDefinition);
        return this;
    }

    @Override
    public final RowDefinition styles(Iterable<String> names, Configurer<CellStyleDefinition> styleDefinition) {
        styles(names);
        style(styleDefinition);
        return this;
    }

    @Override
    public final RowDefinition styles(Iterable<String> styles, Iterable<Configurer<CellStyleDefinition>> styleDefinitions) {
        this.styles(styles);
        for (Configurer<CellStyleDefinition> style : styleDefinitions) {
            this.styleDefinitions.add(style);
        }
        return this;
    }

    @Override
    public final RowDefinition styles(String... names) {
        styles.addAll(Arrays.asList(names));
        return this;
    }

    @Override
    public final RowDefinition styles(Iterable<String> names) {
        for (String name : names) {
            styles.add(name);
        }
        return this;
    }

    AbstractSheetDefinition getSheet() {
        return sheet;
    }

    @Override
    public final RowDefinition group(Configurer<RowDefinition> insideGroupDefinition) {
        createGroup(false, insideGroupDefinition);
        return this;
    }

    @Override
    public final RowDefinition collapse(Configurer<RowDefinition> insideGroupDefinition) {
        createGroup(true, insideGroupDefinition);
        return this;
    }


    private void createGroup(boolean collapsed, Configurer<RowDefinition> insideGroupDefinition) {
        startPositions.add(nextColNumber);
        Configurer.Runner.doConfigure(insideGroupDefinition, this);

        int startPosition = startPositions.remove(startPositions.size() - 1);

        if (nextColNumber - startPosition > 1) {
            int endPosition = nextColNumber - 1;
            doCreateGroup(startPosition, endPosition, collapsed);
        }
    }

    protected abstract void doCreateGroup(int startPosition, int endPosition, boolean collapsed);

    // TODO: make protected
    public CellDefinition getCellByNumber(int oneBasedColumnNumber) {
        return cells.get(oneBasedColumnNumber);
    }

    @Override
    public String toString() {
        return "Row[" + sheet.getName() + "!" + getNumber() + "]";
    }

    protected abstract int getNumber();

    public List<String> getStyles() {
        return Collections.unmodifiableList(styles);
    }
}
