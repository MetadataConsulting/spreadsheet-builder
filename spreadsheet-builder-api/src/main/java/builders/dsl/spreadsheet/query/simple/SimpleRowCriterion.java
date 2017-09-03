package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.impl.Utils;
import builders.dsl.spreadsheet.query.api.CellCriterion;
import builders.dsl.spreadsheet.query.api.RowCriterion;
import builders.dsl.spreadsheet.query.api.Predicate;

final class SimpleRowCriterion extends AbstractCriterion<Cell, RowCriterion> implements RowCriterion {

    private final SimpleSheetCriterion parent;

    SimpleRowCriterion(SimpleSheetCriterion parent) {
        this.parent = parent;
    }

    private SimpleRowCriterion(SimpleSheetCriterion parent, boolean disjoint) {
        super(disjoint);
        this.parent = parent;
    }

    @Override
    public RowCriterion cell(final int from, final int to) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getColumn() >= from && o.getColumn() <= to;
            }
        });
        return this;
    }

    @Override
    public RowCriterion cell(String from, String to) {
        cell(Utils.parseColumn(from), Utils.parseColumn(to));
        return this;
    }

    @Override
    public RowCriterion cell(int from, int to, Configurer<CellCriterion> cellCriterion) {
        cell(from, to);
        cell(cellCriterion);
        return this;
    }

    @Override
    public RowCriterion cell(String from, String to, Configurer<CellCriterion> cellCriterion) {
        cell(from, to);
        cell(cellCriterion);
        return this;
    }

    @Override
    public SimpleRowCriterion cell(final int column) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getColumn() == column;
            }
        });
        return this;
    }

    @Override
    public SimpleRowCriterion cell(final String column) {
        addCondition(new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getColumnAsString().equals(column);
            }
        });
        return this;
    }

    @Override
    public SimpleRowCriterion cell(Configurer<CellCriterion> cellCriterion) {
        SimpleCellCriterion criterion = new SimpleCellCriterion();
        Configurer.Runner.doConfigure(cellCriterion, criterion);
        addCondition(criterion);
        return this;
    }

    @Override
    public SimpleRowCriterion cell(int column, Configurer<CellCriterion> cellCriterion) {
        cell(column);
        cell(cellCriterion);
        return this;
    }

    @Override
    public SimpleRowCriterion cell(String column, Configurer<CellCriterion> cellCriterion) {
        cell(column);
        cell(cellCriterion);
        return this;
    }

    @Override
    public SimpleRowCriterion or(Configurer<RowCriterion> sheetCriterion) {
        return (SimpleRowCriterion) super.or(sheetCriterion);
    }

    @Override
    public RowCriterion having(Predicate<Row> rowPredicate) {
        parent.addCondition(rowPredicate);
        return this;
    }

    @Override
    RowCriterion newDisjointCriterionInstance() {
        return new SimpleRowCriterion(parent, true);
    }
}
