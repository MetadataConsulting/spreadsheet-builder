package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.*;
import org.modelcatalogue.spreadsheet.query.api.PageCriterion;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

public class SimplePageCriterion extends AbstractPageSettingsProvider implements PageCriterion {

    private final SimpleWorkbookCriterion workbookCriterion;

    SimplePageCriterion(SimpleWorkbookCriterion workbookCriterion) {
        this.workbookCriterion = workbookCriterion;
    }

    @Override
    public SimplePageCriterion orientation(final Keywords.Orientation orientation) {
        workbookCriterion.addCondition(new Predicate<Sheet>() {
            @Override
            public boolean test(Sheet o) {
                return orientation.equals(o.getPage().getOrientation());
            }
        });
        return this;
    }

    @Override
    public SimplePageCriterion paper(final Keywords.Paper paper) {
        workbookCriterion.addCondition(new Predicate<Sheet>() {
            @Override
            public boolean test(Sheet o) {
                return paper.equals(o.getPage().getPaper());
            }
        });
        return this;
    }

}
