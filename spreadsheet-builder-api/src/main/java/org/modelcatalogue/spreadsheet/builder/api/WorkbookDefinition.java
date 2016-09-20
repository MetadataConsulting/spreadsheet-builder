package org.modelcatalogue.spreadsheet.builder.api;

import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface WorkbookDefinition extends CanDefineStyle {

    void sheet(String name, @DelegatesTo(SheetDefinition.class) @ClosureParams(value=FromString.class, options = "org.modelcatalogue.spreadsheet.builder.api.SheetDefinition") Closure sheetDefinition);

}
