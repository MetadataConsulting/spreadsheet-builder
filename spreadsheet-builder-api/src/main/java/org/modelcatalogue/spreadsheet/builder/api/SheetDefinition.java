package org.modelcatalogue.spreadsheet.builder.api;


import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.modelcatalogue.spreadsheet.api.Keywords;

public interface SheetDefinition {

    /**
     * Crates new empty row.
     */
    void row();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition closure defining the content of the row
     */
    void row(@DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition closure defining the content of the row
     */
    void row(int row, @DelegatesTo(RowDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.RowDefinition") Closure rowDefinition);

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

    void group(@DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition);
    void collapse(@DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure insideGroupDefinition);

    Object getLocked();

    void password(String password);

    void filter(Keywords.Auto auto);
    Keywords.Auto getAuto();

    /**
     * Configures the basic page settings.
     * @param pageDefinition closure defining the page settings
     */
    void page(@DelegatesTo(PageDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.PageDefinition") Closure pageDefinition);


}
