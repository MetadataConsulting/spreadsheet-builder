package org.modelcatalogue.builder.spreadsheet.api;


import groovy.lang.Closure;
import groovy.lang.DelegatesTo;

public interface Row extends HasStyle {

    void cell();
    void cell(Object value);
    void cell(@DelegatesTo(Cell.class) Closure cellDefinition);
    void cell(int column, @DelegatesTo(Cell.class) Closure cellDefinition);
    void cell(String column, @DelegatesTo(Cell.class) Closure cellDefinition);

    void group(@DelegatesTo(Row.class) Closure insideGroupDefinition);
    void collapse(@DelegatesTo(Row.class) Closure insideGroupDefinition);

}
