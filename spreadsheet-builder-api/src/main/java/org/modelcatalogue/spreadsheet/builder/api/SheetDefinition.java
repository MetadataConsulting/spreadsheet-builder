package org.modelcatalogue.spreadsheet.builder.api;


import org.modelcatalogue.spreadsheet.api.Keywords;

public interface SheetDefinition {

    /**
     * Crates new empty row.
     */
    void row();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition definition of the content of the row
     */
    void row(Configurer<RowDefinition> rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition definition of the content of the row
     */
    void row(int row, Configurer<RowDefinition> rowDefinition);

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

    void group(Configurer<SheetDefinition> insideGroupDefinition);
    void collapse(Configurer<SheetDefinition> insideGroupDefinition);

    Object getLocked();

    void password(String password);

    void filter(Keywords.Auto auto);
    Keywords.Auto getAuto();

    /**
     * Configures the basic page settings.
     * @param pageDefinition definition of the page settings
     */
    void page(Configurer<PageDefinition> pageDefinition);


}
