package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;

public interface RowDefinition extends HasStyle {

    RowDefinition cell();
    RowDefinition cell(Object value);
    RowDefinition cell(Configurer<CellDefinition> cellDefinition);
    RowDefinition cell(int column, Configurer<CellDefinition> cellDefinition);
    RowDefinition cell(String column, Configurer<CellDefinition> cellDefinition);

    RowDefinition group(Configurer<RowDefinition> insideGroupDefinition);
    RowDefinition collapse(Configurer<RowDefinition> insideGroupDefinition);

    RowDefinition style(String name, Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition styles(Iterable<String> names, Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition style(Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition style(String name);
    RowDefinition styles(String... names);
    RowDefinition styles(Iterable<String> names);

}
