package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Comment;
import builders.dsl.spreadsheet.api.Cell;
import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.builder.api.SheetDefinition;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

import java.util.Date;

public interface CellCriterion extends Predicate<Cell> {

    CellCriterion date(Date value);
    CellCriterion date(Predicate<Date> predicate);

    CellCriterion number(Double value);
    CellCriterion number(Predicate<Double> predicate);

    CellCriterion string(String value);
    CellCriterion string(Predicate<String> predicate);

    CellCriterion value(Object value);
    CellCriterion bool(Boolean value);

    CellCriterion style(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleCriterion") Configurer<CellStyleCriterion> styleCriterion);

    CellCriterion rowspan(int span);
    CellCriterion rowspan(Predicate<Integer> predicate);
    CellCriterion colspan(int span);
    CellCriterion colspan(Predicate<Integer> predicate);


    CellCriterion name(String name);
    CellCriterion name(Predicate<String> predicate);

    CellCriterion comment(String comment);
    CellCriterion comment(Predicate<Comment> predicate);

    CellCriterion or(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellCriterion") Configurer<CellCriterion> sheetCriterion);
    CellCriterion having(Predicate<Cell> cellPredicate);

}
