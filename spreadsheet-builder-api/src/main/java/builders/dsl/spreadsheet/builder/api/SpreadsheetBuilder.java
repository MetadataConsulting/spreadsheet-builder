package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;

public interface SpreadsheetBuilder {
    void build(Configurer<WorkbookDefinition> workbookDefinition);
}
