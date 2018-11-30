package builders.dsl.spreadsheet.builder.api;


import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.SheetStateProvider;

import java.util.function.Consumer;

public interface SheetDefinition extends SheetStateProvider {

    /**
     * Crates new empty row.
     */
    SheetDefinition row();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(Consumer<RowDefinition> rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(int row, Consumer<RowDefinition> rowDefinition);

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

    SheetDefinition group(Consumer<SheetDefinition> insideGroupDefinition);
    SheetDefinition collapse(Consumer<SheetDefinition> insideGroupDefinition);

    SheetDefinition state(Keywords.SheetState state);

    SheetDefinition password(String password);

    SheetDefinition filter(Keywords.Auto auto);

    /**
     * Configures the basic page settings.
     * @param pageDefinition definition of the page settings
     */
    SheetDefinition page(Consumer<PageDefinition> pageDefinition);


}
