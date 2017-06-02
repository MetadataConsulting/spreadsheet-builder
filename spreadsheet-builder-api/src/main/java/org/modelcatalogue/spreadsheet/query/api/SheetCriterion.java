package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Page;
import org.modelcatalogue.spreadsheet.api.Row;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

public interface SheetCriterion extends Predicate<Row> {

    Predicate<Row> number(int row);
    Predicate<Row> range(int from, int to);

    void row(Configurer<RowCriterion> rowCriterion);
    void row(int row);
    void row(int row, Configurer<RowCriterion> rowCriterion);
    void row(Predicate<Row> predicate, Configurer<RowCriterion> rowCriterion);
    void row(Predicate<Row> predicate);
    void page(Configurer<PageCriterion> pageCriterion);
    void page(Predicate<Page> predicate);
    void or(Configurer<SheetCriterion> sheetCriterion);
}
