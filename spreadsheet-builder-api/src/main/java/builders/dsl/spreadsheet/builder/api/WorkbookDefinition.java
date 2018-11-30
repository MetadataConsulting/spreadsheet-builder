package builders.dsl.spreadsheet.builder.api;

import java.util.function.Consumer;

public interface WorkbookDefinition extends CanDefineStyle {

    WorkbookDefinition sheet(String name, Consumer<SheetDefinition> sheetDefinition);
    WorkbookDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition);

    WorkbookDefinition apply(Class<? extends Stylesheet> stylesheet);
    WorkbookDefinition apply(Stylesheet stylesheet);

}
