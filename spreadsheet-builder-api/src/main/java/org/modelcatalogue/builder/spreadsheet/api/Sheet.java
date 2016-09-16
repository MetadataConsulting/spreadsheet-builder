package org.modelcatalogue.builder.spreadsheet.api;


import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface Sheet {


    /**
     * @return the parent workbook of the sheet
     */
    Workbook workbook();

    /**
     * @return the name of the sheet
     */
    String name();

    /**
     * Crates new empty row.
     */
    void row ();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition closure defining the content of the row
     */
    void row (@DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition closure defining the content of the row
     */
    void row (int row, @DelegatesTo(Row.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Row") Closure rowDefinition);

    /**
     * Freeze some column or row or both.
     * @param column last freeze column
     * @param row last freeze row
     */
    void freeze(int column, int row);

    /**
     * Freeze some column or row or both.
     * @param column last freeze column
     * @param row last freeze row
     */
    void freeze(String column, int row);

    void group(@DelegatesTo(Sheet.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Sheet") Closure insideGroupDefinition);
    void collapse(@DelegatesTo(Sheet.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.builder.spreadsheet.api.Sheet") Closure insideGroupDefinition);

    Object getLocked();

    void password(String password);

}
