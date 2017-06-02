package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Configurer;
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
    public Predicate<Row> number(final int row) {
        return new Predicate<Row>() {
            @Override
            public boolean test(Row o) {
                return o.getNumber() == row;
            }
        };
    }

    @Override
    public Predicate<Row> range(final int from, final int to) {
        return new Predicate<Row>() {
            @Override
            public boolean test(Row o) {
                return o.getNumber() >= from && o.getNumber() <= to;
            }
        };
    }

    @Override
    public SimpleSheetCriterion row(Configurer<RowCriterion> rowCriterion) {
        SimpleRowCriterion criterion = new SimpleRowCriterion();
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
    public SimpleSheetCriterion row(Predicate<Row> predicate, Configurer<RowCriterion> rowCriterion) {
        row(predicate);
        row(rowCriterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion row(Predicate<Row> predicate) {
        addCondition(predicate);
        return this;
    }

    @Override
    public SimpleSheetCriterion page(Configurer<PageCriterion> pageCriterion) {
        SimplePageCriterion criterion = new SimplePageCriterion(parent);
        Configurer.Runner.doConfigure(pageCriterion, criterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion page(final Predicate<Page> predicate) {
        parent.addCondition(new Predicate<Sheet>() {
            @Override
            public boolean test(Sheet o) {
                return predicate.test(o.getPage());
            }
        });
        return this;
    }

    @Override
    public SimpleSheetCriterion row(int row) {
        row(number(row));
        return this;
    }

    @Override
    public SimpleSheetCriterion or(Configurer<SheetCriterion> sheetCriterion) {
        return (SimpleSheetCriterion) super.or(sheetCriterion);
    }

    Collection<SimpleRowCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }

    @Override
    SheetCriterion newDisjointCriterionInstance() {
        return new SimpleSheetCriterion(true, parent);
    }
}
