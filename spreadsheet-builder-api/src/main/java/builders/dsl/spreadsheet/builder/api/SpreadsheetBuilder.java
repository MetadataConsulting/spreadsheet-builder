package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface SpreadsheetBuilder {
    void build(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkbookDefinition.class) @ClosureParams(value = FromString.class, options = "builders.dsl.spreadsheet.builder.api.WorkbookDefinition") Configurer<WorkbookDefinition> workbookDefinition);
}
