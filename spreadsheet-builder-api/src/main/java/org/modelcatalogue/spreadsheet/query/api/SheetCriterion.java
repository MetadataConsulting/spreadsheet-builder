package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Row;

public interface SheetCriterion {

    Condition<Row> number(int row);

    void row (@DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion);
    void row (int row);
    void row (int row, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion);
    void row (Condition<Row> condition, @DelegatesTo(RowCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowCriterion") Closure rowCriterion);
    void row (Condition<Row> condition);

}
