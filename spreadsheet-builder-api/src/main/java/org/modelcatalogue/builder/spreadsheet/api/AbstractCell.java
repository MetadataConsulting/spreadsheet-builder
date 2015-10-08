package org.modelcatalogue.builder.spreadsheet.api;


public abstract class AbstractCell implements Cell {

    public AutoKeyword getAuto() {
        return AutoKeyword.AUTO;
    }

    public ToKeyword getTo() {
        return ToKeyword.TO;
    }

}
