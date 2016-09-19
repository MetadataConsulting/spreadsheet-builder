package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.query.api.AbstractCriterion;
import org.modelcatalogue.spreadsheet.query.api.Condition;
import org.modelcatalogue.spreadsheet.query.api.SheetCriterion;
import org.modelcatalogue.spreadsheet.query.api.WorkbookCriterion;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;

public final class SimpleWorkbookCriterion extends AbstractCriterion<Sheet> implements WorkbookCriterion {

    private final Collection<SimpleSheetCriterion> criteria = new ArrayList<SimpleSheetCriterion>();

    @Override
    public Condition<Sheet> name(final String name) {
        return new Condition<Sheet>() {
            @Override
            public boolean evaluate(Sheet o) {
                return o.getName().equals(name);
            }
        };
    }

    @Override
    public Condition<Sheet> name(final Condition<String> nameCondition) {
        return new Condition<Sheet>() {
            @Override
            public boolean evaluate(Sheet o) {
                return nameCondition.evaluate(o.getName());
            }
        };
    }

    @Override
    public void sheet(String name) {
        sheet(name(name));
    }

    @Override
    public void sheet(String name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion) {
        sheet(name(name));
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(Condition<Sheet> name) {
        addCondition(name);
    }

    @Override
    public void sheet(Condition<Sheet> name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion) {
        sheet(name);
        sheet(sheetCriterion);
    }

    @Override
    public void sheet(@DelegatesTo(SheetCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion) {
        SimpleSheetCriterion sheet = new SimpleSheetCriterion();
        DefaultGroovyMethods.with(sheet, sheetCriterion);
        criteria.add(sheet);
    }

    public Collection<SimpleSheetCriterion> getCriteria() {
        return Collections.unmodifiableCollection(criteria);
    }
}
