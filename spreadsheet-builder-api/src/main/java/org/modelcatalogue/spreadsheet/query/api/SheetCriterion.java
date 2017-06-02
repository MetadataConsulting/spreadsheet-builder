package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.api.Configurer;

public interface SheetCriterion extends Predicate<Row> {

    Predicate<Row> number(int row);
    Predicate<Row> range(int from, int to);

    SheetCriterion row(Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(int row);
    SheetCriterion row(int row, Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(Predicate<Row> predicate, Configurer<RowCriterion> rowCriterion);
    SheetCriterion row(Predicate<Row> predicate);
    SheetCriterion page(Configurer<PageCriterion> pageCriterion);
    SheetCriterion page(Predicate<Page> predicate);
    SheetCriterion or(Configurer<SheetCriterion> sheetCriterion);
}
