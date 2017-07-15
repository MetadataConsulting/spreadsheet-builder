package org.modelcatalogue.spreadsheet.builder.api;


import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.Keywords;

public interface SheetDefinition {

    /**
     * Crates new empty row.
     */
    SheetDefinition row();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(Configurer<RowDefinition> rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(int row, Configurer<RowDefinition> rowDefinition);

    /**
     * Freeze some column or row or both.
     * @param column last freeze column
     * @param row last freeze row
     */
    SheetDefinition freeze(int column, int row);

    /**
     * Freeze some column or row or both.
     * @param column last freeze column
     * @param row last freeze row
     */
    SheetDefinition freeze(String column, int row);

    SheetDefinition group(Configurer<SheetDefinition> insideGroupDefinition);
    SheetDefinition collapse(Configurer<SheetDefinition> insideGroupDefinition);

    SheetDefinition lock();

    SheetDefinition hide();
    SheetDefinition hideCompletely();
    SheetDefinition show();

    SheetDefinition password(String password);

    SheetDefinition filter(Keywords.Auto auto);

    /**
     * Configures the basic page settings.
     * @param pageDefinition definition of the page settings
     */
    SheetDefinition page(Configurer<PageDefinition> pageDefinition);


}
