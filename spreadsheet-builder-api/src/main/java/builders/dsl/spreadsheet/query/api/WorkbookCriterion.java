package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Sheet;
import builders.dsl.spreadsheet.api.SheetStateProvider;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface WorkbookCriterion extends Predicate<Sheet>, SheetStateProvider {

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetCriterion") Configurer<SheetCriterion> sheetCriterion);
    WorkbookCriterion sheet(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = SheetCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.SheetCriterion") Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = WorkbookCriterion.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.WorkbookCriterion") Configurer<WorkbookCriterion> workbookCriterion);

}
