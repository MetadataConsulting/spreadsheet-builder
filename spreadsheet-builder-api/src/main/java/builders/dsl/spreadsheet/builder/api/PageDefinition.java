package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.PageSettingsProvider;

public interface PageDefinition extends PageSettingsProvider {

    PageDefinition orientation(Keywords.Orientation orientation);
    PageDefinition paper(Keywords.Paper paper);
    FitDimension fit(Keywords.Fit widthOrHeight);

}
