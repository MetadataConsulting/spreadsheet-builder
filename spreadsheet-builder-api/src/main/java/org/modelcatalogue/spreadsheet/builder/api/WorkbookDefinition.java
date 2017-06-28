package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Configurer;

public interface WorkbookDefinition extends CanDefineStyle {

    WorkbookDefinition sheet(String name, Configurer<SheetDefinition> sheetDefinition);

    WorkbookDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition);

    WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet);
    WorkbookDefinition apply(Stylesheet stylesheet);

}
