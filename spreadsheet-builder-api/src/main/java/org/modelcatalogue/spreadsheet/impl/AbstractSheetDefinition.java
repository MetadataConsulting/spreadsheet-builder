package org.modelcatalogue.spreadsheet.impl;

import org.modelcatalogue.spreadsheet.builder.api.SheetDefinition;

public abstract class AbstractSheetDefinition implements SheetDefinition {

       protected abstract void processAutoColumns();
       protected abstract void processAutomaticFilter();

}
