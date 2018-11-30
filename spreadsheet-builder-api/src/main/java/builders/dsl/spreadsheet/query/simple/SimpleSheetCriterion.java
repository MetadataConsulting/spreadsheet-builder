package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.api.Sheet;
import builders.dsl.spreadsheet.query.api.PageCriterion;
import builders.dsl.spreadsheet.query.api.Predicate;
import builders.dsl.spreadsheet.query.api.RowCriterion;
import builders.dsl.spreadsheet.query.api.SheetCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.function.Consumer;

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
    public SimpleSheetCriterion row(Consumer<RowCriterion> rowCriterion) {
        SimpleRowCriterion criterion = new SimpleRowCriterion(this);
        rowCriterion.accept(criterion);
        criteria.add(criterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion row(int row, Consumer<RowCriterion> rowCriterion) {
        row(row);
        row(rowCriterion);
        return this;
    }

    @Override
    public SimpleSheetCriterion page(Consumer<PageCriterion> pageCriterion) {
        SimplePageCriterion criterion = new SimplePageCriterion(parent);
        pageCriterion.accept(criterion);
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
    public SimpleSheetCriterion or(Consumer<SheetCriterion> sheetCriterion) {
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
    public SheetCriterion row(int from, int to, Consumer<RowCriterion> rowCriterion) {
        return null;
    }

    Collection<SimpleRowCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }

    @Override
    SheetCriterion newDisjointCriterionInstance() {
        return new SimpleSheetCriterion(true, parent);
    }

    @Override
    public SheetCriterion state(Keywords.SheetState state) {
        switch (state) {
            case LOCKED:
                parent.addCondition(new Predicate<Sheet>() {
                    @Override
                    public boolean test(Sheet o) {
                        return o.isLocked();
                    }
                });
                return this;
            case VISIBLE:
                parent.addCondition(new Predicate<Sheet>() {
                    @Override
                    public boolean test(Sheet o) {
                        return o.isVisible();
                    }
                });
                return this;
            case HIDDEN:
                parent.addCondition(new Predicate<Sheet>() {
                    @Override
                    public boolean test(Sheet o) {
                        return o.isHidden();
                    }
                });
                return this;
            case VERY_HIDDEN:
                parent.addCondition(new Predicate<Sheet>() {
                    @Override
                    public boolean test(Sheet o) {
                        return o.isVeryHidden();
                    }
                });
                return this;
        }
        throw new IllegalStateException("Unknown sheet state: " + state);
    }
}
