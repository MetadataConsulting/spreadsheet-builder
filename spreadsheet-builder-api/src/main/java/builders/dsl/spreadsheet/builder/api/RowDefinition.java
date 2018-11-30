package builders.dsl.spreadsheet.builder.api;

import java.util.function.Consumer;

public interface RowDefinition extends HasStyle {

    RowDefinition cell();
    RowDefinition cell(Object value);
    RowDefinition cell(Consumer<CellDefinition> cellDefinition);
    RowDefinition cell(int column, Consumer<CellDefinition> cellDefinition);
    RowDefinition cell(String column, Consumer<CellDefinition> cellDefinition);

    RowDefinition group(Consumer<RowDefinition> insideGroupDefinition);
    RowDefinition collapse(Consumer<RowDefinition> insideGroupDefinition);

    RowDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition styles(Iterable<String> names, Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition style(Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition style(String name);
    RowDefinition styles(String... names);
    RowDefinition styles(Iterable<String> names);

}
