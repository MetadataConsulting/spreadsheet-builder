package org.modelcatalogue.spreadsheet.builder.api;

import org.modelcatalogue.spreadsheet.api.Configurer;
import org.modelcatalogue.spreadsheet.api.SheetStateProvider;

public interface WorkbookDefinition extends CanDefineStyle, SheetStateProvider {

    WorkbookDefinition sheet(String name, Configurer<SheetDefinition> sheetDefinition);

    WorkbookDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition);

    WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet);
    WorkbookDefinition apply(Stylesheet stylesheet);

}
