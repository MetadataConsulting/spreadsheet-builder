package org.modelcatalogue.spreadsheet.query.api;

public interface Condition<T> {

    /**
     * @param o object to be evaluated
     * @return <code>true</code> if the object passes the condition
     */
    boolean evaluate(T o);

    class EqualsTo<T> implements Condition<T> {
        private final Object other;

        public EqualsTo(Object other) {
            this.other = other;
        }

        @Override
        public boolean evaluate(T o) {
            return other.equals(o);
        }
    }

    class Negation<T> implements Condition<T> {
        private final Condition<T> otherCondition;

        public Negation(Condition<T> otherCondition) {
            this.otherCondition = otherCondition;
        }

        @Override
        public boolean evaluate(T o) {
            return !otherCondition.evaluate(o);
        }
    }

    class Match implements Condition<String> {
        private final String regexp;

        public Match(String regexp) {
            this.regexp = regexp;
        }

        @Override
        public boolean evaluate(String o) {
            return o.matches(regexp);
        }
    }
}
