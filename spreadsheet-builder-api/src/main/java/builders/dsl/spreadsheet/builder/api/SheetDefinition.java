package builders.dsl.spreadsheet.builder.api;


import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Keywords;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface SheetDefinition {

    /**
     * Crates new empty row.
     */
    SheetDefinition row();

    /**
     * Creates new row in the spreadsheet.
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowDefinition") Configurer<RowDefinition> rowDefinition);

    /**
     * Creates new row in the spreadsheet.
     * @param row row number (1 based - the same as is shown in the file)
     * @param rowDefinition definition of the content of the row
     */
    SheetDefinition row(int row, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowDefinition") Configurer<RowDefinition> rowDefinition);

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

    SheetDefinition group(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetDefinition") Configurer<SheetDefinition> insideGroupDefinition);
    SheetDefinition collapse(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetDefinition") Configurer<SheetDefinition> insideGroupDefinition);

    SheetDefinition state(Keywords.SheetState state);

    SheetDefinition password(String password);

    SheetDefinition filter(Keywords.Auto auto);

    /**
     * Configures the basic page settings.
     * @param pageDefinition definition of the page settings
     */
    SheetDefinition page(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = PageDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.PageDefinition") Configurer<PageDefinition> pageDefinition);


}
