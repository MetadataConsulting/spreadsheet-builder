package org.modelcatalogue.spreadsheet.builder.api;


public abstract class AbstractCellDefinition implements CellDefinition {

    public Keywords.Auto getAuto() {
        return Keywords.Auto.AUTO;
    }

    public Keywords.To getTo() {
        return Keywords.To.TO;
    }

    public Keywords.Image getImage() {
        return Keywords.Image.IMAGE;
    }

}
