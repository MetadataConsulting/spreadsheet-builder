package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Cell;

public interface RowCriterion {

    Condition<Cell> column(int number);
    Condition<Cell> columnAsString(String name);

    void cell(Condition<Cell> condition);

    void cell(@DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);
    void cell(int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);
    void cell(String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);

}
