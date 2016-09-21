package org.modelcatalogue.spreadsheet.query.simple;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;
import org.codehaus.groovy.runtime.DefaultGroovyMethods;
import org.modelcatalogue.spreadsheet.query.api.Predicate;

import java.util.ArrayList;
import java.util.List;

abstract class AbstractCriterion<T> implements Predicate<T> {

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

    abstract Predicate<T> newDisjointCriterionInstance();

    public void or(Closure sheetCriterion) {
        Predicate<T> criterion = newDisjointCriterionInstance();
        DefaultGroovyMethods.with(criterion, sheetCriterion);
        addCondition(criterion);
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
