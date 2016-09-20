package org.modelcatalogue.spreadsheet.api;

import java.util.Collection;

public interface Row {

    int getNumber();
    Sheet getSheet();

    Collection<? extends Cell> getCells();

    Row above();
    Row above(int howMany);
    Row bellow();
    Row bellow(int howMany);

}
