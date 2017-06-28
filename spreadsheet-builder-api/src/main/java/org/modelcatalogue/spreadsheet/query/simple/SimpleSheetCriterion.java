package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.PageCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.RowCriterion;
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

final class SimpleSheetCriterion extends AbstractCriterion<Row, SheetCriterion> implements SheetCriterion {

    private final Collection<SimpleRowCriterion> criteria = new ArrayList<SimpleRowCriterion>();
    private final SimpleWorkbookCriterion parent;

    SimpleSheetCriterion(SimpleWorkbookCriterion parent) {
        this.parent = parent;
    }

    private SimpleSheetCriterion(boolean disjoint, SimpleWorkbookCriterion parent) {
        super(disjoint);
        this.parent = parent;
    }

    @Override
    public SimpleSheetCriterion row(Configurer<RowCriterion> rowCriterion) {
        SimpleRowCriterion criterion = new SimpleRowCriterion(this);
        Configurer.Runner.doConfigure(rowCriterion, criterion);
        criteria.add(criterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion row(int row, Configurer<RowCriterion> rowCriterion) {
        row(row);
        row(rowCriterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion page(Configurer<PageCriterion> pageCriterion) {
        SimplePageCriterion criterion = new SimplePageCriterion(parent);
        Configurer.Runner.doConfigure(pageCriterion, criterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion row(final int row) {
        addCondition(new Predicate<Row>() {
            @Override
            public boolean test(Row o) {
                return o.getNumber() == row;
            }
        });
        return this;
    }

    @Override
    public SimpleSheetCriterion or(Configurer<SheetCriterion> sheetCriterion) {
        return (SimpleSheetCriterion) super.or(sheetCriterion);
    }

    @Override
    public SheetCriterion having(Predicate<Sheet> sheetPredicate) {
        parent.addCondition(sheetPredicate);
        return this;
    }

    @Override
    public SheetCriterion row(final int from, final int to) {
        addCondition(new Predicate<Row>() {
            @Override
            public boolean test(Row o) {
                return o.getNumber() >= from && o.getNumber() <= to;
            }
        });
        return this;
    }

    @Override
    public SheetCriterion row(int from, int to, Configurer<RowCriterion> rowCriterion) {
        return null;
    }

    Collection<SimpleRowCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }

    @Override
    SheetCriterion newDisjointCriterionInstance() {
        return new SimpleSheetCriterion(true, parent);
    }
}
