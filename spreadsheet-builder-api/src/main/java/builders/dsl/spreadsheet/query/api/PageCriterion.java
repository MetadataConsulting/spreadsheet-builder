package builders.dsl.spreadsheet.query.api;

import builders.dsl.spreadsheet.api.Keywords;
import builders.dsl.spreadsheet.api.Page;
import builders.dsl.spreadsheet.api.PageSettingsProvider;

import java.util.function.Predicate;

public interface PageCriterion extends PageSettingsProvider {

    PageCriterion orientation(Keywords.Orientation orientation);
    PageCriterion paper(Keywords.Paper paper);
    PageCriterion having(Predicate<Page> pagePredicate);
}
