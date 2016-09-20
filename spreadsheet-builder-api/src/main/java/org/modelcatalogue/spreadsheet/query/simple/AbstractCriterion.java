package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.query.api.Predicate;

import java.util.ArrayList;
import java.util.List;

public abstract class AbstractCriterion<C> {

    private final List<Predicate<C>> predicates = new ArrayList<Predicate<C>>();

    void addCondition(Predicate<C> predicate) {
        predicates.add(predicate);
    }

    public boolean passesAnyCondition(C object) {
        if (predicates.isEmpty()) {
            return true;
        }
        for (Predicate<C> predicate : predicates) {
            if (predicate.test(object)) {
                return true;
            }
        }
        return false;
    }

}
