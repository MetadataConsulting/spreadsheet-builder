package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface Workbook {

    /**
     * Declare a named style.
     * @param name name of the style
     * @param styleDefinition definition of the style
     */
    void style(String name, @DelegatesTo(CellStyle.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.CellStyle") Closure styleDefinition);
    void sheet(String name, @DelegatesTo(Sheet.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Sheet") Closure sheetDefinition);
    

}
