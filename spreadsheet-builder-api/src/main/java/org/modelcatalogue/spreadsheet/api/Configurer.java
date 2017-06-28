package org.modelcatalogue.spreadsheet.api;

import space.jasan.support.groovy.closure.GroovyClosure;

/**
 * Interface similar to Consumer which will attempt to set the delegate if executed with Groovy Closure
 * using {@link Runner#doConfigure(Configurer, Object)}
 */
public interface Configurer<T> {

    /**
     * @param t object to be configured
     * @deprecated do not use this method directly, use {@link Runner#doConfigure(Configurer, Object)} instead
     */
    @Deprecated void configure(T t);

    class Noop<U> implements Configurer<U> {
        public static final Configurer NOOP = new Noop();
        @Override public void configure(U u) {}
    }

    class Runner {
        public static <T> T doConfigure(Configurer<T> configurer, T target) {
            if (configurer == null) {
                return target;
            }
            GroovyClosure.setDelegate(configurer, target);
            configurer.configure(target);
            return target;
        }
    }

}
