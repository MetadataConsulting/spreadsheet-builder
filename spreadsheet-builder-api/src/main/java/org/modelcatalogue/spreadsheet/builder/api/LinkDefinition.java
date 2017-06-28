package org.modelcatalogue.spreadsheet.builder.api;

import java.util.Map;

public interface LinkDefinition {

    CellDefinition name(String name);

    CellDefinition email(String email);
    CellDefinition email(Map<String, ?> parameters, String email);

    CellDefinition url(String url);

    CellDefinition file(String path);
}
