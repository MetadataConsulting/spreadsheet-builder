package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Sheet;
import org.modelcatalogue.spreadsheet.builder.api.Configurer;

public interface WorkbookCriterion extends Predicate<Sheet> {

    Predicate<Sheet> name(String name);
    Predicate<Sheet> name(Predicate<String> namePredicate);

    void sheet(String name);
    void sheet(String name, Configurer<SheetCriterion> sheetCriterion);

    void sheet(Predicate<Sheet> predicate);
    void sheet(Predicate<Sheet> predicate, Configurer<SheetCriterion> sheetCriterion);

    void sheet(Configurer<SheetCriterion> sheetCriterion);

    void or(Configurer<WorkbookCriterion> workbookCriterion);

}
