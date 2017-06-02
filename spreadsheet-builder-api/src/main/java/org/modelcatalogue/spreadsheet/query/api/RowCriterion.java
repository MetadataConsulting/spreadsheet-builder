package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

public interface RowCriterion extends Predicate<Cell> {

    Predicate<Cell> column(int number);
    Predicate<Cell> columnAsString(String name);
    Predicate<Cell> range(int from, int to);
    Predicate<Cell> range(String from, String to);

    void cell(Predicate<Cell> predicate);
    void cell(int column);
    void cell(String column);

    void cell(Configurer<CellCriterion> cellCriterion);
    void cell(int column, Configurer<CellCriterion> cellCriterion);
    void cell(String column, Configurer<CellCriterion> cellCriterion);
    void cell(Predicate<Cell> predicate, Configurer<CellCriterion> cellCriterion);
    void or(Configurer<RowCriterion> rowCriterion);
}
