package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;

public interface Workbook {

    /**
     * Declare a named style.
     * @param name name of the style
     * @param styleDefinition definition of the style
     */
    void style(String name, @DelegatesTo(CellStyle.class) Closure styleDefinition);
    void sheet(String name, @DelegatesTo(Sheet.class) Closure sheetDefinition);
    

}
