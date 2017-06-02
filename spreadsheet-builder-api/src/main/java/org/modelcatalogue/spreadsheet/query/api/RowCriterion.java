package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Configurer;

public interface RowCriterion extends Predicate<Cell> {

    Predicate<Cell> column(int number);
    Predicate<Cell> columnAsString(String name);
    Predicate<Cell> range(int from, int to);
    Predicate<Cell> range(String from, String to);

    RowCriterion cell(Predicate<Cell> predicate);
    RowCriterion cell(int column);
    RowCriterion cell(String column);

    RowCriterion cell(Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(int column, Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(String column, Configurer<CellCriterion> cellCriterion);
    RowCriterion cell(Predicate<Cell> predicate, Configurer<CellCriterion> cellCriterion);
    RowCriterion or(Configurer<RowCriterion> rowCriterion);
}
