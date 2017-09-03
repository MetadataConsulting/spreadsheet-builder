package builders.dsl.spreadsheet.impl;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.CellStyleDefinition;
import builders.dsl.spreadsheet.builder.api.RowDefinition;

import java.util.*;

public abstract class AbstractRowDefinition implements RowDefinition {

    protected final AbstractSheetDefinition sheet;

    private List<String> styles = new ArrayList<String>();
    private List<Configurer<CellStyleDefinition>> styleDefinitions = new ArrayList<Configurer<CellStyleDefinition>>();

    private final List<Integer> startPositions = new ArrayList<Integer>();
    private int nextColNumber;
    private final Map<Integer, AbstractCellDefinition> cells = new LinkedHashMap<Integer, AbstractCellDefinition>();

    protected AbstractRowDefinition(AbstractSheetDefinition sheet) {
        this.sheet = sheet;
    }

    private AbstractCellDefinition findOrCreateCell(int zeroBasedCellNumber) {
        AbstractCellDefinition cell = cells.get(zeroBasedCellNumber + 1);

        if (cell != null) {
            return cell;
        }

        cell = createCell(zeroBasedCellNumber);

        cells.put(zeroBasedCellNumber + 1, cell);

        return cell;
    }

    protected abstract AbstractCellDefinition createCell(int zeroBasedCellNumber);

    @Override
    public final RowDefinition cell() {
        cell(null);
        return this;
    }

    @Override
    public final RowDefinition cell(Object value) {
        AbstractCellDefinition cell = findOrCreateCell(nextColNumber++);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            cell.styles(styles, styleDefinitions);
        }

        cell.value(value);

        cell.resolve();

        return this;
    }

    @Override
    public final RowDefinition cell(Configurer<CellDefinition> cellDefinition) {
        AbstractCellDefinition poiCell = findOrCreateCell(nextColNumber);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            poiCell.styles(styles, styleDefinitions);
        }

        Configurer.Runner.doConfigure(cellDefinition, poiCell);

        nextColNumber += poiCell.getColspan();

        handleSpans(poiCell);

        poiCell.resolve();

        return this;
    }

    protected abstract void handleSpans(AbstractCellDefinition poiCell);

    @Override
    public final RowDefinition cell(int column, Configurer<CellDefinition> cellDefinition) {
        AbstractCellDefinition poiCell = findOrCreateCell(column - 1);

        if (!styles.isEmpty() || !styleDefinitions.isEmpty()) {
            poiCell.styles(styles, styleDefinitions);
        }

        Configurer.Runner.doConfigure(cellDefinition, poiCell);

        nextColNumber = column - 1 + poiCell.getColspan();

        handleSpans(poiCell);

        poiCell.resolve();

        return this;
    }

    @Override
    public final RowDefinition cell(String column, Configurer<CellDefinition> cellDefinition) {
        cell(Utils.parseColumn(column), cellDefinition);
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

    @Override
    public String toString() {
        return "Row[" + sheet.getName() + "!" + getNumber() + "]";
    }

    protected abstract int getNumber();

    List<String> getStyles() {
        return Collections.unmodifiableList(styles);
    }
}
