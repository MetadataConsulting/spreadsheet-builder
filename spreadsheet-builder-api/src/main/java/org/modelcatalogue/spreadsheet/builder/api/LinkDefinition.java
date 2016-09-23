package org.modelcatalogue.spreadsheet.builder.api;

import java.util.Map;

public interface LinkDefinition {

    void name(String name);

    void email(String email);
    void email(Map<String, ?> parameters, String email);

    void url(String url);

    void file(String path);
}
