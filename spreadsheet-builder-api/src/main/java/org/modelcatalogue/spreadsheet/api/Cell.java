package org.modelcatalogue.spreadsheet.api;

public interface Cell {

    int getColumn();
    <T> T read(Class<T> type);
    Row getRow();

}
