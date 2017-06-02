package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Configurer;

public interface WorkbookCriterion extends Predicate<Sheet> {

    Predicate<Sheet> name(String name);
    Predicate<Sheet> name(Predicate<String> namePredicate);

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion sheet(Predicate<Sheet> predicate);
    WorkbookCriterion sheet(Predicate<Sheet> predicate, Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion sheet(Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(Configurer<WorkbookCriterion> workbookCriterion);

}
