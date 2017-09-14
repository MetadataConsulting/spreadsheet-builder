package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Sheet;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface SheetCriterion extends Predicate<Row> {

    SheetCriterion row(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowCriterion") Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int row);
    SheetCriterion row(int from, int to);
    SheetCriterion row(int row, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowCriterion") Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int from, int to, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowCriterion") Configurer<RowCriterion> rowCriterion);
    SheetCriterion page(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = PageCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.PageCriterion") Configurer<PageCriterion> pageCriterion);
    SheetCriterion or(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetCriterion") Configurer<SheetCriterion> sheetCriterion);
    SheetCriterion having(Predicate<Sheet> sheetPredicate);
    SheetCriterion state(Keywords.SheetState state);
}
