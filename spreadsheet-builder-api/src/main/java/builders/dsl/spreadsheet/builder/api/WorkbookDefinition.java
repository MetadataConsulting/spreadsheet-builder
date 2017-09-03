package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;
import builders.dsl.spreadsheet.api.SheetStateProvider;

public interface WorkbookDefinition extends CanDefineStyle, SheetStateProvider {

    WorkbookDefinition sheet(String name, Configurer<SheetDefinition> sheetDefinition);

    WorkbookDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition);

    WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet);
    WorkbookDefinition apply(Stylesheet stylesheet);

}
