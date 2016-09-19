package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Sheet;

public interface WorkbookCriterion {

    Condition<Sheet> name(String name);
    Condition<Sheet> name(Condition<String> nameCondition);

    void sheet(String name);
    void sheet(String name, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion);

    void sheet(Condition<Sheet> condition);
    void sheet(Condition<Sheet> condition, @DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion);

    void sheet(@DelegatesTo(SheetCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.SheetCriterion") Closure sheetCriterion);

}
