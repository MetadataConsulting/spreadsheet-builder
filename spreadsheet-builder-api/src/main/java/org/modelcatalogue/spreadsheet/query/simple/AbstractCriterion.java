package org.modelcatalogue.spreadsheet.query.simple;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

import java.util.ArrayList;
import java.util.List;

abstract class AbstractCriterion<T, C extends Predicate<T>> implements Predicate<T> {

    private final List<Predicate<T>> predicates = new ArrayList<Predicate<T>>();
    private final boolean disjoint;

    AbstractCriterion() {
        this(false);
    }

    AbstractCriterion(boolean disjoint) {
        this.disjoint = disjoint;
    }

    @Override
    public boolean test(T o) {
        if (disjoint) {
            return passesAnyCondition(o);
        }
        return passesAllConditions(o);
    }

    abstract C newDisjointCriterionInstance();

    public AbstractCriterion<T, C> or(Configurer<C> sheetCriterion) {
        C criterion = newDisjointCriterionInstance();
        Configurer.Runner.doConfigure(sheetCriterion, criterion);
        addCondition(criterion);
        return this;
    }

    void addCondition(Predicate<T> predicate) {
        predicates.add(predicate);
    }

    private boolean passesAnyCondition(T object) {
        if (predicates.isEmpty()) {
            return true;
        }
        for (Predicate<T> predicate : predicates) {
            if (predicate.test(object)) {
                return true;
            }
        }
        return false;
    }

    private boolean passesAllConditions(T object) {
        if (predicates.isEmpty()) {
            return true;
        }
        for (Predicate<T> predicate : predicates) {
            if (!predicate.test(object)) {
                return false;
            }
        }
        return true;
    }

}
