package org.modelcatalogue.builder.spreadsheet.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;

public interface HasStyle {
    void style(@DelegatesTo(CellStyle.class) Closure styleDefinition);
    void style(String name);
}
