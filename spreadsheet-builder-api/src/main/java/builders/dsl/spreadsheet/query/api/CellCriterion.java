package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Comment;

import java.util.Date;
import java.util.function.Consumer;

public interface CellCriterion extends Predicate<Cell> {

    CellCriterion date(Date value);
    CellCriterion date(Predicate<Date> predicate);

    CellCriterion number(Double value);
    CellCriterion number(Predicate<Double> predicate);

    CellCriterion string(String value);
    CellCriterion string(Predicate<String> predicate);

    CellCriterion value(Object value);
    CellCriterion bool(Boolean value);

    CellCriterion style(Consumer<CellStyleCriterion> styleCriterion);

    CellCriterion rowspan(int span);
    CellCriterion rowspan(Predicate<Integer> predicate);
    CellCriterion colspan(int span);
    CellCriterion colspan(Predicate<Integer> predicate);


    CellCriterion name(String name);
    CellCriterion name(Predicate<String> predicate);

    CellCriterion comment(String comment);
    CellCriterion comment(Predicate<Comment> predicate);

    CellCriterion or(Consumer<CellCriterion> sheetCriterion);
    CellCriterion having(Predicate<Cell> cellPredicate);

}
