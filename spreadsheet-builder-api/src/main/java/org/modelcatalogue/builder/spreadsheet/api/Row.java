package org.modelcatalogue.builder.spreadsheet.api;


import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface Row extends HasStyle {

    void cell();
    void cell(Object value);
    void cell(@DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition);
    void cell(int column, @DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition);
    void cell(String column, @DelegatesTo(Cell.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Cell") Closure cellDefinition);

    void group(@DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure insideGroupDefinition);
    void collapse(@DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure insideGroupDefinition);

}
