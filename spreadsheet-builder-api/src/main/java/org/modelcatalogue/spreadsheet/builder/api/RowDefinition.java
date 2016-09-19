package org.modelcatalogue.spreadsheet.builder.api;


import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface RowDefinition extends HasStyle {

    void cell();
    void cell(Object value);
    void cell(@DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellDefinition") Closure cellDefinition);
    void cell(int column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellDefinition") Closure cellDefinition);
    void cell(String column, @DelegatesTo(CellDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellDefinition") Closure cellDefinition);

    void group(@DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowDefinition") Closure insideGroupDefinition);
    void collapse(@DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.RowDefinition") Closure insideGroupDefinition);

}
