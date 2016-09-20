package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Cell;

public interface RowCriterion {

    Predicate<Cell> column(int number);
    Predicate<Cell> columnAsString(String name);

    void cell(Predicate<Cell> predicate);

    void cell(@DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);
    void cell(int column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);
    void cell(String column, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);
    void cell(Predicate<Cell> predicate, @DelegatesTo(CellCriterion.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellCriterion") Closure cellCriterion);

}
