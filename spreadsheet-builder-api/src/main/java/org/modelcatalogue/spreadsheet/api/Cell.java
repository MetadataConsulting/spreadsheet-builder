package org.modelcatalogue.spreadsheet.api;

public interface Cell extends Spannable {

    int getColumn();
    Object getValue();
    String getColumnAsString();
    <T> T read(Class<T> type);
    Row getRow();

    String getName();
    Comment getComment();
    CellStyle getStyle();

    Cell getAbove();
    Cell getBellow();
    Cell getLeft();
    Cell getRight();
    Cell getAboveLeft();
    Cell getAboveRight();
    Cell getBellowLeft();
    Cell getBellowRight();

    class Util {

        private Util() {}

    }

}
