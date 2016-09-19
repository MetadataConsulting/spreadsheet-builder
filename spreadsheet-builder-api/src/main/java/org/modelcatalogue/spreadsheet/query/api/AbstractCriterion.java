package org.modelcatalogue.spreadsheet.query.api;

import java.util.ArrayList;
import java.util.List;

public abstract class AbstractCriterion<C> {

    private final List<Condition<C>> conditions = new ArrayList<Condition<C>>();

    protected void addCondition(Condition<C> condition) {
        conditions.add(condition);
    }

    public boolean passesAnyCondition(C object) {
        if (conditions.isEmpty()) {
            return true;
        }
        for (Condition<C> condition : conditions) {
            if (condition.evaluate(object)) {
                return true;
            }
        }
        return false;
    }

}
