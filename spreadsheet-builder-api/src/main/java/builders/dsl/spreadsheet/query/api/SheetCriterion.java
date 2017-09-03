package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Sheet;

public interface SheetCriterion extends Predicate<Row> {

    SheetCriterion row(Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int row);
    SheetCriterion row(int from, int to);
    SheetCriterion row(int row, Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int from, int to, Configurer<RowCriterion> rowCriterion);
    SheetCriterion page(Configurer<PageCriterion> pageCriterion);
    SheetCriterion or(Configurer<SheetCriterion> sheetCriterion);
    SheetCriterion having(Predicate<Sheet> sheetPredicate);
    SheetCriterion state(Keywords.SheetState state);
}
