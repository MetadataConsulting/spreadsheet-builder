package org.modelcatalogue.spreadsheet.builder.api;

import space.jasan.support.groovy.closure.GroovyClosure;

/**
 * Single abstract method type which similar to Consumer which guarantees that it will attempt to set the delegate
 * if executed with Groovy Closure.
 */
public abstract class Configurer<T> {

    public final void configure(T t) {
        GroovyClosure.setDelegate(this, t);
        doConfigure(t);
    }

    protected abstract void doConfigure(T configurable);

    public static final <T> Configurer<T> noop() {
        return (Configurer<T>) Noop.NOOP;
    }

    private static class Noop<U> extends Configurer<U> {
        static final Configurer NOOP = new Noop();

        @Override public void doConfigure(U u) {}
    }

}
