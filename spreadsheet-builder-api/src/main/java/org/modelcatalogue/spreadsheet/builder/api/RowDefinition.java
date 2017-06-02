package org.modelcatalogue.spreadsheet.builder.api;

public interface RowDefinition extends HasStyle {

    void cell();
    void cell(Object value);
    void cell(Configurer<CellDefinition> cellDefinition);
    void cell(int column, Configurer<CellDefinition> cellDefinition);
    void cell(String column, Configurer<CellDefinition> cellDefinition);

    void group(Configurer<RowDefinition> insideGroupDefinition);
    void collapse(Configurer<RowDefinition> insideGroupDefinition);

}
