package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.query.api.CellCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.RowCriterion;

final class SimpleRowCriterion extends AbstractCriterion<Cell> implements RowCriterion {

    SimpleRowCriterion() {}

    private SimpleRowCriterion(boolean disjoint) {
        super(disjoint);
    }

    @Override
    public Predicate<Cell> column(final int number) {
        return new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return number == o.getColumn();
            }
        };
    }

    @Override
    public Predicate<Cell> columnAsString(final String name) {
        return new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return name.equals(o.getColumnAsString());
            }
        };
    }

    @Override
    public Predicate<Cell> range(final int from, final int to) {
        return new Predicate<Cell>() {
            @Override
            public boolean test(Cell o) {
                return o.getColumn() >= from && o.getColumn() <= to;
            }
        };
    }

    @Override
    public Predicate<Cell> range(String from, String to) {
        return range(Cell.Util.parseColumn(from), Cell.Util.parseColumn(to));
    }

    @Override
    public void cell(Predicate<Cell> predicate) {
        addCondition(predicate);
    }

    @Override
    public void cell(int column) {
        cell(column(column));
    }

    @Override
    public void cell(String column) {
        cell(columnAsString(column));
    }

    @Override
    public void cell(@DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        SimpleCellCriterion criterion = new SimpleCellCriterion();
        DefaultGroovyMethods.with(criterion, cellCriterion);
        addCondition(criterion);
    }

    @Override
    public void cell(int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        addCondition(column(column));
        cell(cellCriterion);
    }

    @Override
    public void cell(String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        addCondition(columnAsString(column));
        cell(cellCriterion);
    }

    @Override
    public void cell(Predicate<Cell> predicate, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellCriterion") Closure cellCriterion) {
        addCondition(predicate);
        cell(cellCriterion);
    }

    @Override
    Predicate<Cell> newDisjointCriterionInstance() {
        return new SimpleRowCriterion(true);
    }
}
