package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.api.PageSettingsProvider;

public interface PageCriterion extends PageSettingsProvider {

    PageCriterion orientation(Keywords.Orientation orientation);
    PageCriterion paper(Keywords.Paper paper);

}
