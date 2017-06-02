package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Cell;
import org.modelcatalogue.spreadsheet.api.Comment;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

import java.util.Date;

public interface CellCriterion extends Predicate<Cell> {

    void date(Date value);
    void date(Predicate<Date> predicate);

    void number(Double value);
    void number(Predicate<Double> predicate);

    void string(String value);
    void string(Predicate<String> predicate);

    void value(Object value);
    void bool(Boolean value);

    void style(Configurer<CellStyleCriterion> styleCriterion);

    void rowspan(int span);
    void rowspan(Predicate<Integer> predicate);
    void colspan(int span);
    void colspan(Predicate<Integer> predicate);


    void name(String name);
    void name(Predicate<String> predicate);

    void comment(String comment);
    void comment(Predicate<Comment> predicate);

    void or(Configurer<CellCriterion> sheetCriterion);

}
