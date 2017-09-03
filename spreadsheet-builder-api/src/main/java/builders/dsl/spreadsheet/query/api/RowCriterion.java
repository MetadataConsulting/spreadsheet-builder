package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Row;

public interface RowCriterion extends Predicate<Cell> {

    RowCriterion cell(int column);
    RowCriterion cell(int from, int to);
    RowCriterion cell(String column);
    RowCriterion cell(String from, String to);

    RowCriterion cell(Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(int column, Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(int from, int to, Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(String column, Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(String from, String to, Configurer<CellCriterion> cellCriterion);
    RowCriterion or(Configurer<RowCriterion> rowCriterion);
    RowCriterion having(Predicate<Row> rowPredicate);
}
