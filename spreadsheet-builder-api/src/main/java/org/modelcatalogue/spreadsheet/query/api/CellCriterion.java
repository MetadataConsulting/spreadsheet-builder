package org.modelcatalogue.spreadsheet.query.api;

import java.util.Date;

public interface CellCriterion {

    void date(Date value);
    void date(Condition<Date> condition);

    void number(Double value);
    void number(Condition<Double> condition);

    void string(String value);
    void string(Condition<String> condition);

    void value(Object value);
    void bool(Boolean value);

//    void rowspan(int span);
//    void colspan(int span);
//    void name(String name);
//    void comment(String comment);
//    void style(StyleCriterion comment);

}
