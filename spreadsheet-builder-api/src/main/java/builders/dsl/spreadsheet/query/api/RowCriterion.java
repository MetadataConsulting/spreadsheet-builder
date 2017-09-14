package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Row;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface RowCriterion extends Predicate<Cell> {

    RowCriterion cell(int column);
    RowCriterion cell(int from, int to);
    RowCriterion cell(String column);
    RowCriterion cell(String from, String to);

    RowCriterion cell(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(int column, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(int from, int to, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(String column, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(String from, String to, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> cellCriterion);
    RowCriterion or(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowCriterion") Configurer<RowCriterion> rowCriterion);
    RowCriterion having(Predicate<Row> rowPredicate);
}
