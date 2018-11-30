package builders.dsl.spreadsheet.query.simple;

import builders.dsl.spreadsheet.query.api.Predicate;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

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

    public AbstractCriterion<T, C> or(Consumer<C> sheetCriterion) {
        C criterion = newDisjointCriterionInstance();
        sheetCriterion.accept(criterion);
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
