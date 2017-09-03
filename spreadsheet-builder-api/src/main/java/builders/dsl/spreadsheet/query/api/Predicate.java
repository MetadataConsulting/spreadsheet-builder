package builders.dsl.spreadsheet.query.api;

public interface Predicate<T> {

    /**
     * @param o object to be evaluated
     * @return <code>true</code> if the object passes the condition
     */
    boolean test(T o);

    class EqualsTo<T> implements Predicate<T> {
        private final Object other;

        public EqualsTo(Object other) {
            this.other = other;
        }

        @Override
        public boolean test(T o) {
            return other.equals(o);
        }
    }

    class Negation<T> implements Predicate<T> {
        private final Predicate<T> otherPredicate;

        public Negation(Predicate<T> otherPredicate) {
            this.otherPredicate = otherPredicate;
        }

        @Override
        public boolean test(T o) {
            return !otherPredicate.test(o);
        }
    }

    class Match implements Predicate<String> {
        private final String regexp;

        public Match(String regexp) {
            this.regexp = regexp;
        }

        @Override
        public boolean test(String o) {
            return o.matches(regexp);
        }
    }
}
