package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.Sheet;
import builders.dsl.spreadsheet.api.SheetStateProvider;

public interface WorkbookCriterion extends Predicate<Sheet>, SheetStateProvider {

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, Configurer<SheetCriterion> sheetCriterion);
    WorkbookCriterion sheet(Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(Configurer<WorkbookCriterion> workbookCriterion);

}
