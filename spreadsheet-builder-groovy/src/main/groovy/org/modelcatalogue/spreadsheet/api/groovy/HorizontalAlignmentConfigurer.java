package org.modelcatalogue.spreadsheet.api.groovy;


import org.modelcatalogue.spreadsheet.builder.api.CellStyleDefinition;

public interface HorizontalAlignmentConfigurer {

    // following methods must return something otherwise they are not considered to be getters

    CellStyleDefinition getRight();
    CellStyleDefinition getLeft();

    CellStyleDefinition getGeneral();
    CellStyleDefinition getCenter();
    CellStyleDefinition getFill();
    CellStyleDefinition getJustify();
    CellStyleDefinition getCenterSelection();

}
