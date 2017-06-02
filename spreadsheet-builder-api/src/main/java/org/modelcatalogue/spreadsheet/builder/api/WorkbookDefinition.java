package org.modelcatalogue.spreadsheet.builder.api;

public interface WorkbookDefinition extends CanDefineStyle {

    void sheet(String name, Configurer<SheetDefinition> sheetDefinition);

}
