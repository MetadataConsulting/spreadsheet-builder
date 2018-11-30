package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Row;
import builders.dsl.spreadsheet.api.Sheet;
import builders.dsl.spreadsheet.api.SheetStateProvider;

import java.util.function.Consumer;
import java.util.function.Predicate;

public interface SheetCriterion extends Predicate<Row>, SheetStateProvider {

    SheetCriterion row(Consumer<RowCriterion> rowCriterion);
    SheetCriterion row(int row);
    SheetCriterion row(int from, int to);
    SheetCriterion row(int row, Consumer<RowCriterion> rowCriterion);
    SheetCriterion row(int from, int to, Consumer<RowCriterion> rowCriterion);
    SheetCriterion page(Consumer<PageCriterion> pageCriterion);
    SheetCriterion or(Consumer<SheetCriterion> sheetCriterion);
    SheetCriterion having(Predicate<Sheet> sheetPredicate);
    SheetCriterion state(Keywords.SheetState state);

}
