package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.query.api.Condition;
import org.modelcatalogue.spreadsheet.query.api.RowCriterion;
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public final class SimpleSheetCriterion extends AbstractCriterion<Row> implements SheetCriterion {

    private final Collection<SimpleRowCriterion> criteria = new ArrayList<SimpleRowCriterion>();

    @Override
    public Condition<Row> number(final int row) {
        return new Condition<Row>() {
            @Override
            public boolean evaluate(Row o) {
                return o.getNumber() == row;
            }
        };
    }

    @Override
    public void row(@DelegatesTo(RowCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion) {
        SimpleRowCriterion criterion = new SimpleRowCriterion();
        DefaultGroovyMethods.with(criterion, rowCriterion);
        criteria.add(criterion);
    }

    @Override
    public void row(int row, @DelegatesTo(RowCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion) {
        row(row);
        row(rowCriterion);
    }

    @Override
    public void row(Condition<Row> condition, @DelegatesTo(RowCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion) {
        row(condition);
        row(rowCriterion);
    }

    @Override
    public void row(Condition<Row> condition) {
        addCondition(condition);
    }

    @Override
    public void row(int row) {
        row(number(row));
    }

    public Collection<SimpleRowCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }
}
