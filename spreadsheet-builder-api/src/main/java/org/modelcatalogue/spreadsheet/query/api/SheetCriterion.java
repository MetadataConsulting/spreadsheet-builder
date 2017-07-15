package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.Sheet;

public interface SheetCriterion extends Predicate<Row> {

    SheetCriterion row(Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int row);
    SheetCriterion row(int from, int to);
    SheetCriterion row(int row, Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int from, int to, Configurer<RowCriterion> rowCriterion);
    SheetCriterion page(Configurer<PageCriterion> pageCriterion);
    SheetCriterion or(Configurer<SheetCriterion> sheetCriterion);
    SheetCriterion having(Predicate<Sheet> sheetPredicate);
    SheetCriterion hide();
    SheetCriterion hideCompletely();
    SheetCriterion show();
}
