package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.query.api.AbstractCriterion;
import org.modelcatalogue.spreadsheet.query.api.CellCriterion;
import org.modelcatalogue.spreadsheet.query.api.Condition;
import org.modelcatalogue.spreadsheet.query.api.RowCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public class SimpleRowCriterion extends AbstractCriterion<Cell> implements RowCriterion {

    private final Collection<SimpleCellCriterion> criteria = new ArrayList<SimpleCellCriterion>();

    @Override
    public Condition<Cell> column(final int number) {
        return new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                return number == o.getColumn();
            }
        };
    }

    @Override
    public Condition<Cell> columnAsString(final String name) {
        return new Condition<Cell>() {
            @Override
            public boolean evaluate(Cell o) {
                return name.equals(o.getColumnAsString());
            }
        };
    }

    @Override
    public void cell(Condition<Cell> condition) {
        addCondition(condition);
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

    public Collection<SimpleCellCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }
}
