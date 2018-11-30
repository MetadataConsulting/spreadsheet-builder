package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Sheet;

import java.util.function.Consumer;

public interface WorkbookCriterion extends Predicate<Sheet> {

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, Consumer<SheetCriterion> sheetCriterion);
    WorkbookCriterion sheet(Consumer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(Consumer<WorkbookCriterion> workbookCriterion);

}
