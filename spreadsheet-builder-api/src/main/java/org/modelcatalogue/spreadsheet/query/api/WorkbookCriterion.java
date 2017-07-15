package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.SheetStateProvider;

public interface WorkbookCriterion extends Predicate<Sheet>, SheetStateProvider {

    WorkbookCriterion sheet(String name);
    WorkbookCriterion sheet(String name, Configurer<SheetCriterion> sheetCriterion);
    WorkbookCriterion sheet(Configurer<SheetCriterion> sheetCriterion);

    WorkbookCriterion or(Configurer<WorkbookCriterion> workbookCriterion);

}
