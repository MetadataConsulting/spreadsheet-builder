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

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public final class SimpleRowCriterion extends AbstractCriterion<Cell> implements RowCriterion {

    private final Collection<SimpleCellCriterion> criteria = new ArrayList<SimpleCellCriterion>();

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
    public void cell(Predicate<Cell> predicate) {
        addCondition(predicate);
    }

    @Override
    public void cell(@DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion) {
        SimpleCellCriterion criterion = new SimpleCellCriterion();
        DefaultGroovyMethods.with(criterion, cellCriterion);
        criteria.add(criterion);
    }

    @Override
    public void cell(int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion) {
        addCondition(column(column));
        cell(cellCriterion);
    }

    @Override
    public void cell(String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion) {
        addCondition(columnAsString(column));
        cell(cellCriterion);
    }

    @Override
    public void cell(Predicate<Cell> predicate, @DelegatesTo(CellCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion) {
        addCondition(predicate);
        cell(cellCriterion);
    }

    public Collection<SimpleCellCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }
}
