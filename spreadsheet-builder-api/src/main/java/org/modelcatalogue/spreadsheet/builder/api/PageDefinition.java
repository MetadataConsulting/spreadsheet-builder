package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Keywords;
import org.modelcatalogue.spreadsheet.api.PageSettingsProvider;

public interface PageDefinition extends PageSettingsProvider {

    void orientation(Keywords.Orientation orientation);
    void paper(Keywords.Paper paper);
    FitDimension fit(Keywords.Fit widthOrHeight);

}
