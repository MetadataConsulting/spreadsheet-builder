package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

final class SimpleWorkbookCriterion extends AbstractCriterion<Sheet, WorkbookCriterion> implements WorkbookCriterion {

    private final Collection<SimpleSheetCriterion> criteria = new ArrayList<SimpleSheetCriterion>();

    private SimpleWorkbookCriterion(boolean disjoint) {
        super(disjoint);
    }

    SimpleWorkbookCriterion() {}

    @Override
    public Predicate<Sheet> name(final String name) {
        return new Predicate<Sheet>() {
            @Override
            public boolean test(Sheet o) {
                return o.getName().equals(name);
            }
        };
    }

    @Override
    public Predicate<Sheet> name(final Predicate<String> namePredicate) {
        return new Predicate<Sheet>() {
            @Override
            public boolean test(Sheet o) {
                return namePredicate.test(o.getName());
            }
        };
    }

    @Override
    public void sheet(String name) {
        sheet(name(name));
    }

    @Override
    public void sheet(String name, Configurer<SheetCriterion> sheetCriterion) {
        sheet(name(name));
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(Predicate<Sheet> name) {
        addCondition(name);
    }

    @Override
    public void sheet(Predicate<Sheet> name, Configurer<SheetCriterion> sheetCriterion) {
        sheet(name);
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(Configurer<SheetCriterion> sheetCriterion) {
        SimpleSheetCriterion sheet = new SimpleSheetCriterion(this);
        Configurer.Runner.doConfigure(sheetCriterion, sheet);
        criteria.add(sheet);
    }

    Collection<SimpleSheetCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }

    @Override
    WorkbookCriterion newDisjointCriterionInstance() {
        return new SimpleWorkbookCriterion(true);
    }
}
