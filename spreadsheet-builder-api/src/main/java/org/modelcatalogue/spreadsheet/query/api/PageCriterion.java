package org.modelcatalogue.spreadsheet.query.api;

import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.api.PageSettingsProvider;

public interface PageCriterion extends PageSettingsProvider {

    void orientation(Keywords.Orientation orientation);
    void paper(Keywords.Paper paper);

}
