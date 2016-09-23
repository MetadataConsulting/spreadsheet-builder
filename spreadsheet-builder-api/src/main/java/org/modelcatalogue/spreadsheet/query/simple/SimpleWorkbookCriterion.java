package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.query.api.Predicate;
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

final class SimpleWorkbookCriterion extends AbstractCriterion<Sheet> implements WorkbookCriterion {

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
    public void sheet(String name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        sheet(name(name));
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(Predicate<Sheet> name) {
        addCondition(name);
    }

    @Override
    public void sheet(Predicate<Sheet> name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        sheet(name);
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(@DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.SheetCriterion") Closure sheetCriterion) {
        SimpleSheetCriterion sheet = new SimpleSheetCriterion();
        DefaultGroovyMethods.with(sheet, sheetCriterion);
        criteria.add(sheet);
    }

    Collection<SimpleSheetCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }

    @Override
    Predicate<Sheet> newDisjointCriterionInstance() {
        return new SimpleWorkbookCriterion(true);
    }
}
