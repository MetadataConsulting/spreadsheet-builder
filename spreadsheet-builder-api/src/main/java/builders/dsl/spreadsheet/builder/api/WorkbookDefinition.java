package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.SheetStateProvider;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface WorkbookDefinition extends CanDefineStyle, SheetStateProvider {

    WorkbookDefinition sheet(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetDefinition") Configurer<SheetDefinition> sheetDefinition);
    WorkbookDefinition style(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);

    WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet);
    WorkbookDefinition apply(Stylesheet stylesheet);

}
