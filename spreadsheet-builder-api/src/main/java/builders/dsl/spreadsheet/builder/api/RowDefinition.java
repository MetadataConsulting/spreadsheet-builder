package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.Configurer;
import groovy.lang.Closure;
import groovy.lang.DelegatesTo;
import groovy.transform.stc.ClosureParams;
import groovy.transform.stc.FromString;

public interface RowDefinition extends HasStyle {

    RowDefinition cell();
    RowDefinition cell(Object value);
    RowDefinition cell(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellDefinition") Configurer<CellDefinition> cellDefinition);
    RowDefinition cell(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellDefinition") Closure cellDefinition);
    RowDefinition cell(int column, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellDefinition") Configurer<CellDefinition> cellDefinition);
    RowDefinition cell(String column, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellDefinition") Configurer<CellDefinition> cellDefinition);

    RowDefinition group(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowDefinition") Configurer<RowDefinition> insideGroupDefinition);
    RowDefinition collapse(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = RowDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.RowDefinition") Configurer<RowDefinition> insideGroupDefinition);

    RowDefinition style(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition styles(Iterable<String> names, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition style(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellStyleDefinition.class) @ClosureParams(value=FromString.class, options = "builders.dsl.spreadsheet.builder.api.CellStyleDefinition") Configurer<CellStyleDefinition> styleDefinition);
    RowDefinition style(String name);
    RowDefinition styles(String... names);
    RowDefinition styles(Iterable<String> names);

}
