package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Configurer;

public interface WorkbookCriterion extends Predicate<Sheet> {

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, Configurer<SheetCriterion> sheetCriterion);
    WorkbookCriterion sheet(Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(Configurer<WorkbookCriterion> workbookCriterion);

}
