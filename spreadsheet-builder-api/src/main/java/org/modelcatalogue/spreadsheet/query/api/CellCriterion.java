package org.modelcatalogue.spreadsheet.query.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Comment;

import java.util.Date;

public interface CellCriterion {

    void date(Date value);
    void date(Predicate<Date> predicate);

    void number(Double value);
    void number(Predicate<Double> predicate);

    void string(String value);
    void string(Predicate<String> predicate);

    void value(Object value);
    void bool(Boolean value);

    void style(@DelegatesTo(CellStyleCriterion.class) @ClosureParams(value = FromString.class, options = "org.modelcatalogue.spreadsheet.query.api.CellStyleCriterion") Closure styleCriterion);

    void rowspan(int span);
    void rowspan(Predicate<Integer> predicate);
    void colspan(int span);
    void colspan(Predicate<Integer> predicate);


    void name(String name);
    void name(Predicate<String> predicate);

    void comment(String comment);
    void comment(Predicate<Comment> predicate);

}
