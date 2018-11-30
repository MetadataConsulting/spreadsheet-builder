package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Row;

import java.util.function.Consumer;

public interface RowCriterion extends Predicate<Cell> {

    RowCriterion cell(int column);
    RowCriterion cell(int from, int to);
    RowCriterion cell(String column);
    RowCriterion cell(String from, String to);

    RowCriterion cell(Consumer<CellCriterion> cellCriterion);
    RowCriterion cell(int column, Consumer<CellCriterion> cellCriterion);
    RowCriterion cell(int from, int to, Consumer<CellCriterion> cellCriterion);
    RowCriterion cell(String column, Consumer<CellCriterion> cellCriterion);
    RowCriterion cell(String from, String to, Consumer<CellCriterion> cellCriterion);
    RowCriterion or(Consumer<RowCriterion> rowCriterion);
    RowCriterion having(Predicate<Row> rowPredicate);
}
